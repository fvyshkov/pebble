/**
 * AG Grid-based pivot view (Tree Data edition).
 *
 * Mirrors the row hierarchy produced by the legacy PivotGrid: rows within an
 * analytic are nested by parent_id, and when a leaf of analytic N has a
 * sub-analytic N+1 it becomes a group row whose children are that analytic's
 * rows. We feed AG Grid a flat list of rows with explicit `path: string[]`
 * and `treeData={true}`.
 */
import { useEffect, useMemo, useState, useCallback, useRef } from 'react'
import { Box, Chip, CircularProgress, Dialog, DialogContent, DialogTitle, LinearProgress, Snackbar, Tooltip, Typography, Button as MUIButton, IconButton } from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import * as am5 from '@amcharts/amcharts5'
import * as am5xy from '@amcharts/amcharts5/xy'
import * as am5percent from '@amcharts/amcharts5/percent'
import am5themes_Animated from '@amcharts/amcharts5/themes/Animated'
import { AgGridReact } from 'ag-grid-react'
import {
  ModuleRegistry,
  AllCommunityModule,
  themeAlpine,
  type ColDef,
  type ColGroupDef,
  type CellValueChangedEvent,
  type GridApi,
  type GridReadyEvent,
  type CellKeyDownEvent,
} from 'ag-grid-community'

// Variant of themeAlpine with visible vertical + horizontal cell borders.
// Explicit theme params override the default `none` for column/row borders.
const themeAlpineWithBorders = themeAlpine.withParams({
  columnBorder: { style: 'solid', width: 1, color: '#e0e0e0' },
  rowBorder: { style: 'solid', width: 1, color: '#f0f0f0' },
  headerColumnBorder: { style: 'solid', width: 1, color: '#d5d5d5' },
})
import {
  AllEnterpriseModule,
  LicenseManager,
} from 'ag-grid-enterprise'
import * as api from '../../api'
import type { SheetAnalytic, Analytic, AnalyticRecord, CellData } from '../../types'
import FormulaEditor from './FormulaEditor'
import MoreVertOutlined from '@mui/icons-material/MoreVertOutlined'

ModuleRegistry.registerModules([AllCommunityModule, AllEnterpriseModule])
const licenseKey = (import.meta as any).env?.VITE_AG_GRID_LICENSE
if (licenseKey) LicenseManager.setLicenseKey(licenseKey)

type CellRule = 'manual' | 'sum_children' | 'formula' | 'empty'

// ── Local tree helpers (mirrored from PivotGrid) ──────────────────────────
interface RecordNode {
  record: AnalyticRecord
  data: Record<string, any>
  children: RecordNode[]
}
function buildRecordTree(records: AnalyticRecord[]): RecordNode[] {
  const byParent: Record<string, AnalyticRecord[]> = { root: [] }
  for (const r of records) (byParent[r.parent_id || 'root'] ||= []).push(r)
  const build = (pid: string | null): RecordNode[] =>
    (byParent[pid || 'root'] || []).map(r => {
      const data = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
      return { record: r, data, children: build(r.id) }
    })
  return build(null)
}
function getLeaves(nodes: RecordNode[]): RecordNode[] {
  const out: RecordNode[] = []
  const walk = (ns: RecordNode[]) => {
    for (const n of ns) n.children.length === 0 ? out.push(n) : walk(n.children)
  }
  walk(nodes)
  return out
}
function recordLabel(n: RecordNode): string {
  return (n.data && n.data.name) || n.record.id.slice(0, 8)
}

/** Period level rank: M=0, Q=1, H=2, Y=3. Returns -1 if unknown. */
function periodLevelRank(pk: string): number {
  if (/^\d{4}-M\d{2}$/.test(pk) || /^\d{4}-\d{2}$/.test(pk)) return 0  // M
  if (/^\d{4}-Q\d$/.test(pk)) return 1  // Q
  if (/^\d{4}-H\d$/.test(pk)) return 2  // H
  if (/^\d{4}-Y$/.test(pk)) return 3    // Y
  return -1
}
const LEVEL_RANK: Record<string, number> = { M: 0, Q: 1, H: 2, Y: 3 }

/** Filter record tree: keep only records at minLevel or above (coarser). */
function filterRecordsByLevel(nodes: RecordNode[], minLevel: string): RecordNode[] {
  const minRank = LEVEL_RANK[minLevel] ?? 0
  const walk = (ns: RecordNode[]): RecordNode[] => {
    const out: RecordNode[] = []
    for (const n of ns) {
      const pk: string = n.data?.period_key || ''
      const rank = periodLevelRank(pk)
      if (rank >= 0 && rank < minRank) {
        // This record is below min level — skip it entirely
        continue
      }
      // Keep this node, but filter its children too
      out.push({ ...n, children: walk(n.children) })
    }
    return out
  }
  return walk(nodes)
}

/** Filter flat records list by period level. */
function filterFlatRecordsByLevel(recs: AnalyticRecord[], minLevel: string): AnalyticRecord[] {
  const minRank = LEVEL_RANK[minLevel] ?? 0
  return recs.filter(r => {
    const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
    const pk: string = d?.period_key || ''
    const rank = periodLevelRank(pk)
    return rank < 0 || rank >= minRank
  })
}

// ── Props ──────────────────────────────────────────────────────────────────
interface CalcProgress {
  done: number
  total: number
  sheet?: string
  computed?: number
  totalCells?: number
  startedAt?: number
}

interface Props {
  sheetId: string
  modelId: string
  currentUserId?: string
  /** Model-wide recalc progress (while a calculate-model/stream is in flight). */
  calcProgress?: CalcProgress | null
  /** 'data' — values; 'formulas' — show rule label or formula text (admin). */
  mode?: 'data' | 'formulas'
}

interface RowDatum {
  /** Hierarchy path of unique record IDs (+ suffix) used by AG Grid tree data. */
  path: string[]
  /** Display label for this node. */
  label: string
  /** True iff this row has no sub-analytic children AND no parent record
   *  children — only leaf rows carry cell values. */
  isLeaf: boolean
  /** Per-analytic record ID map so we can build coord keys for each column. */
  recordIds: Record<string, string>
  /** unit string from record data (shown on group rows). */
  unit?: string
  /** Column leaf (period) values: p_<periodRecId> → number | '' */
  [key: string]: any
}

// ── External recalc store (AG Grid instantiates status panels once,
//    so they can't receive React state via props — we subscribe to a tiny
//    pub/sub store instead). ────────────────────────────────────────────────
const recalcStore = {
  running: false,
  listeners: new Set<() => void>(),
  set(v: boolean) {
    if (this.running === v) return
    this.running = v
    this.listeners.forEach(l => l())
  },
  subscribe(l: () => void) {
    this.listeners.add(l)
    return () => { this.listeners.delete(l) }
  },
}

// Custom AG Grid status panel: flowing (indeterminate) progress bar shown
// in place of the default "Rows" panel while a recalc is in flight.
function RecalcStatusPanel() {
  const [, tick] = useState(0)
  useEffect(() => recalcStore.subscribe(() => tick(n => n + 1)), [])
  if (!recalcStore.running) {
    return <Box sx={{ minWidth: 180 }} />
  }
  return (
    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, minWidth: 180, px: 1 }}>
      <LinearProgress sx={{ flex: 1, height: 4, borderRadius: 2 }} />
      <span style={{ fontSize: 12, color: '#1976d2' }}>пересчёт</span>
    </Box>
  )
}

// ── Selection-stats store (Excel-style status bar) ──────────────────────────
type SelectionStats = {
  count: number        // total cells in selection
  numCount: number     // numeric cells
  sum: number
  avg: number
  min: number | null
  max: number | null
}
const selectionStore = {
  stats: { count: 0, numCount: 0, sum: 0, avg: 0, min: null, max: null } as SelectionStats,
  listeners: new Set<() => void>(),
  set(s: SelectionStats) { this.stats = s; this.listeners.forEach(l => l()) },
  subscribe(l: () => void) { this.listeners.add(l); return () => { this.listeners.delete(l) } },
}

function SelectionStatusPanel() {
  const [, tick] = useState(0)
  useEffect(() => selectionStore.subscribe(() => tick(n => n + 1)), [])
  const s = selectionStore.stats
  const fmt = (n: number) => n.toLocaleString('ru-RU', { maximumFractionDigits: 2 })
  const hasNums = s.numCount > 0
  const hasSel = s.count > 0
  return (
    <Box sx={{
      display: 'flex', alignItems: 'center', gap: 2, px: 1.5, minHeight: 28,
      fontSize: 12, color: '#555', fontVariantNumeric: 'tabular-nums',
    }}>
      {hasSel ? (
        <>
          <span>Выделено: <b>{s.count}</b></span>
          <span>Чисел: <b>{s.numCount}</b></span>
          <span>Сумма: <b>{hasNums ? fmt(s.sum) : '—'}</b></span>
          <span>Среднее: <b>{hasNums ? fmt(s.avg) : '—'}</b></span>
          <span>Мин: <b>{hasNums && s.min != null ? fmt(s.min) : '—'}</b></span>
          <span>Макс: <b>{hasNums && s.max != null ? fmt(s.max) : '—'}</b></span>
        </>
      ) : (
        <span style={{ color: '#999' }}>
          Выделите ячейки — в статус-баре появятся сумма, среднее, мин и макс
        </span>
      )}
    </Box>
  )
}

// ── Formula tooltip: green on black ─
function FormulaTooltip(props: any) {
  const { value } = props
  if (!value) return null
  return (
    <div style={{
      background: '#1a1a1a',
      color: '#4caf50',
      padding: '8px 12px',
      borderRadius: 6,
      fontSize: 13,
      fontFamily: 'monospace',
      maxWidth: 400,
      wordBreak: 'break-word',
      whiteSpace: 'pre-wrap',
      boxShadow: '0 4px 12px rgba(0,0,0,0.4)',
    }}>
      {value}
    </div>
  )
}

// ── Small recalc progress indicator (spinner + tooltip, no extra section) ─
function CalcProgressChip({ calcProgress, localRunning }: {
  calcProgress?: CalcProgress | null
  localRunning: boolean
}) {
  // Tick every second so the elapsed counter in the tooltip stays live.
  const [, forceTick] = useState(0)
  useEffect(() => {
    if (!calcProgress?.startedAt) return
    const id = window.setInterval(() => forceTick(n => n + 1), 1000)
    return () => window.clearInterval(id)
  }, [calcProgress?.startedAt])

  const elapsedMs = calcProgress?.startedAt ? Date.now() - calcProgress.startedAt : 0
  const elapsedStr = (() => {
    const s = Math.floor(elapsedMs / 1000)
    const m = Math.floor(s / 60)
    return m > 0 ? `${m} мин ${s % 60} с` : `${s} с`
  })()

  const computed = calcProgress?.computed ?? 0
  const totalCells = calcProgress?.totalCells ?? 0
  const cellsLabel = totalCells > 0
    ? `${computed} / ${totalCells} клеток`
    : `${computed} клеток`
  const sheetsLabel = calcProgress
    ? `листов ${calcProgress.done} из ${calcProgress.total}`
    : ''

  const tooltipLines = [
    calcProgress ? `Идёт ${elapsedStr}` : 'Локальный пересчёт…',
    calcProgress ? cellsLabel : '',
    calcProgress ? sheetsLabel : '',
    calcProgress?.sheet ? `Сейчас: ${calcProgress.sheet}` : '',
  ].filter(Boolean).join('\n')

  return (
    <Tooltip
      title={<span style={{ whiteSpace: 'pre-line' }}>{tooltipLines}</span>}
      placement="top"
      arrow
    >
      <Box sx={{ display: 'inline-flex', alignItems: 'center', gap: 0.5, cursor: 'default' }}>
        <CircularProgress size={11} />
        {calcProgress && totalCells > 0 && (
          <span style={{ color: '#1976d2', fontVariantNumeric: 'tabular-nums' }}>
            {computed}/{totalCells}
          </span>
        )}
        {calcProgress && totalCells === 0 && (
          <span style={{ color: '#1976d2' }}>пересчёт…</span>
        )}
        {!calcProgress && localRunning && (
          <span style={{ color: '#1976d2' }}>сохраняю…</span>
        )}
      </Box>
    </Tooltip>
  )
}

// ── Component ──────────────────────────────────────────────────────────────
export default function PivotGridAG({ sheetId, modelId, currentUserId, calcProgress, mode = 'data' }: Props) {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [rowData, setRowData] = useState<RowDatum[]>([])
  const [columnDefs, setColumnDefs] = useState<(ColDef | ColGroupDef)[]>([])
  const [sheetName, setSheetName] = useState('')
  const [recalcRunning, setRecalcRunning] = useState(false)
  // Mirror local/external recalc state into the store that feeds the custom
  // AG Grid status panel (which lives outside the React tree).
  useEffect(() => {
    recalcStore.set(recalcRunning || !!calcProgress)
  }, [recalcRunning, calcProgress])
  // Analytic-pinning state (drag row → toolbar to fix an analytic on one value)
  const [pinned, setPinned] = useState<Record<string, string>>({})
  // Period-level aggregate toggles (Year / Quarter …).
  // 3 states: 'hidden' | 'start' | 'end' — controls whether sum columns
  // for that level appear before leaves, after leaves, or not at all.
  type ColLevelState = 'hidden' | 'start' | 'end'
  const [colLevelToggles, setColLevelToggles] = useState<Record<number, ColLevelState>>({})
  const [colLevelNames, setColLevelNames] = useState<{ level: number; label: string }[]>([])
  const colTreeRef = useRef<RecordNode[]>([])
  const colLevelTogglesRef = useRef<Record<number, ColLevelState>>({})
  // colId → leafIds it sums. Used to flash Σ-columns whenever one of their
  // underlying leaf period cells changes.
  const sumColLeavesRef = useRef<Record<string, string[]>>({})
  useEffect(() => { colLevelTogglesRef.current = colLevelToggles }, [colLevelToggles])
  const [analyticsMap, setAnalyticsMap] = useState<Record<string, Analytic>>({})
  const [recordNames, setRecordNames] = useState<Record<string, string>>({})
  const vsLoadedRef = useRef(false)
  // Persisted AG Grid column state (widths, order, pinned, hidden) — declared
  // here so the debounced save effect can depend on `columnStateVersion`.
  const savedColumnStateRef = useRef<any[] | null>(null)
  const [columnStateVersion, setColumnStateVersion] = useState(0)

  // Formula editor dialog state (opened from hover ⋮ button on a cell).
  const [formulaEditorOpen, setFormulaEditorOpen] = useState(false)
  const [formulaEditorKey, setFormulaEditorKey] = useState('')
  // Snackbar offered after saving a per-cell formula — one-click "save as
  // rule on the indicator" (promote-cell API). Hidden if we can't derive the
  // main analytic / indicator id from coord.
  const [promoteSnack, setPromoteSnack] = useState<
    { coordKey: string; formula: string; indicatorId: string } | null
  >(null)
  // Cached per-load: main analytic id + its index in dbOrd (so we can extract
  // indicator_id from a coord_key by splitting on '|' and picking that slot).
  const mainAnalyticIdRef = useRef<string | null>(null)
  const mainAnalyticIdxRef = useRef<number>(-1)

  // Chart overlay + history dialog
  const [chartOverlay, setChartOverlay] = useState<{ type: string; labels: string[]; datasets: { label: string; data: number[] }[] } | null>(null)
  const chartDivRef = useRef<HTMLDivElement>(null)
  const amRootRef = useRef<am5.Root | null>(null)
  const [historyOpen, setHistoryOpen] = useState(false)
  const [historyData, setHistoryData] = useState<any[]>([])
  const [historyKey, setHistoryKey] = useState('')

  // Cmd/Ctrl+R — Excel fill-right over the selected range. Handled at the
  // window level (capture phase) so we beat the browser's reload shortcut
  // AND we don't rely on AG Grid's onCellKeyDown firing for this key.
  // Same for Cmd/Ctrl+D (fill-down) — consolidated here for consistency.
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      if (!(e.ctrlKey || e.metaKey)) return
      const k = e.key.toLowerCase()
      if (k !== 'r' && k !== 'd') return
      const api = gridApiRef.current
      if (!api) return
      e.preventDefault()
      e.stopPropagation()
      const ranges = api.getCellRanges()
      if (!ranges || ranges.length === 0) return
      const range = ranges[0]
      const cols = range.columns
      if (!cols || cols.length === 0) return
      const si = range.startRow?.rowIndex ?? 0
      const ei = range.endRow?.rowIndex ?? si
      const rMin = Math.min(si, ei), rMax = Math.max(si, ei)
      const isPeriodCol = (id: string) => id.startsWith('p_')
      const ruleOf = (data: any, colId: string) =>
        (data?.[`__rule_${colId.slice(2)}`] as CellRule) || 'manual'
      if (k === 'd') {
        const source = api.getDisplayedRowAtIndex(rMin)
        if (!source?.data) return
        for (const col of cols) {
          const colId = col.getColId()
          if (!isPeriodCol(colId)) continue
          const val = source.data[colId]
          for (let r = rMin + 1; r <= rMax; r++) {
            const node = api.getDisplayedRowAtIndex(r)
            if (!node?.data || !node.data.isLeaf) continue
            if (ruleOf(node.data, colId) !== 'manual') continue
            node.setDataValue(colId, val)
          }
        }
      } else {
        for (let r = rMin; r <= rMax; r++) {
          const node = api.getDisplayedRowAtIndex(r)
          if (!node?.data || !node.data.isLeaf) continue
          const firstCol = cols[0].getColId()
          if (!isPeriodCol(firstCol)) continue
          const val = node.data[firstCol]
          for (let c = 1; c < cols.length; c++) {
            const colId = cols[c].getColId()
            if (!isPeriodCol(colId)) continue
            if (ruleOf(node.data, colId) !== 'manual') continue
            node.setDataValue(colId, val)
          }
        }
      }
    }
    window.addEventListener('keydown', h, { capture: true })
    return () => window.removeEventListener('keydown', h, { capture: true } as any)
  }, [])

  // Arrow keys anywhere on the page → focus the grid and start navigating,
  // so the user doesn't have to click into the grid first.
  useEffect(() => {
    const h = (e: KeyboardEvent) => {
      const k = e.key
      if (k !== 'ArrowUp' && k !== 'ArrowDown' && k !== 'ArrowLeft' && k !== 'ArrowRight') return
      if (e.ctrlKey || e.metaKey || e.altKey) return
      // Skip if the user is typing into an input / textarea / contenteditable.
      const t = e.target as HTMLElement | null
      if (t) {
        const tag = t.tagName
        if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return
        if (t.isContentEditable) return
      }
      const api = gridApiRef.current
      if (!api) return
      // If the grid already owns a focused cell, let AG Grid handle it natively.
      if (api.getFocusedCell()) return
      // Find first leaf row + first period column and focus it.
      const total = api.getDisplayedRowCount()
      if (total === 0) return
      let firstLeaf = -1
      for (let i = 0; i < total; i++) {
        const n = api.getDisplayedRowAtIndex(i)
        if (n?.data?.isLeaf) { firstLeaf = i; break }
      }
      if (firstLeaf < 0) firstLeaf = 0
      const cols = api.getAllDisplayedColumns()
      const firstPeriod = cols.find(c => c.getColId().startsWith('p_'))
      if (!firstPeriod) return
      e.preventDefault()
      e.stopPropagation()
      api.setFocusedCell(firstLeaf, firstPeriod.getColId())
      api.ensureIndexVisible(firstLeaf)
      api.ensureColumnVisible(firstPeriod.getColId())
    }
    window.addEventListener('keydown', h)
    return () => window.removeEventListener('keydown', h)
  }, [])

  // Refs (stable across re-renders, avoid stale closure in editors/handlers)
  const gridApiRef = useRef<GridApi | null>(null)
  const cellMapRef = useRef<Record<string, string>>({})
  const cellRuleRef = useRef<Record<string, CellRule>>({})
  /** coord_key → formula text (only populated for `formula`-rule cells). */
  const formulaMapRef = useRef<Record<string, string>>({})
  /** coord_key → {formula, source, kind} from POST /cells/resolved-formulas.
   *  Populated only in mode='formulas' to show indicator-rule formulas on
   *  cells that don't have a per-cell formula. */
  const resolvedFormulaMapRef = useRef<Record<string, { formula: string; source: string; kind: string }>>({})
  const colLeafIdsRef = useRef<string[]>([])
  const dbOrdRef = useRef<string[]>([])
  const colAIdRef = useRef<string>('')
  const pinnedRef = useRef<Record<string, string>>({})
  const analyticsRef = useRef<Record<string, Analytic>>({})
  /** For each row analytic, the single root record ID (if exactly one root with children exists). */
  const rootRecordByAidRef = useRef<Record<string, string>>({})

  // Build coord key for (row × period) using DB order and right-truncation
  // lookup against cellMapRef.
  const lookupCell = useCallback((parts: string[]): { value: string; rule: CellRule; key: string } => {
    for (let n = parts.length; n >= 2; n--) {
      const k = parts.slice(0, n).join('|')
      if (cellMapRef.current[k] !== undefined) {
        return { value: cellMapRef.current[k], rule: cellRuleRef.current[k] || 'manual', key: k }
      }
    }
    return { value: '', rule: 'manual', key: parts.join('|') }
  }, [])

  const load = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      const sa: SheetAnalytic[] = await api.listSheetAnalytics(sheetId)
      if (sa.length < 2) throw new Error('Нужно минимум 2 аналитики на листе (колонки + строки)')

      const aMap: Record<string, Analytic> = {}
      const rMap: Record<string, RecordNode[]> = {}
      const recNameMap: Record<string, string> = {}
      // Build min_period_level lookup from bindings
      const periodLevelByAid: Record<string, string | null> = {}
      for (const b of sa) periodLevelByAid[b.analytic_id] = b.min_period_level ?? null
      await Promise.all(sa.map(async b => {
        const [analytic, recs] = await Promise.all([
          api.getAnalytic(b.analytic_id),
          api.listRecords(b.analytic_id),
        ])
        aMap[b.analytic_id] = analytic
        const minLvl = periodLevelByAid[b.analytic_id]
        const filteredRecs = minLvl ? filterFlatRecordsByLevel(recs, minLvl) : recs
        rMap[b.analytic_id] = buildRecordTree(filteredRecs)
        for (const r of filteredRecs) {
          const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
          recNameMap[r.id] = (d && d.name) || r.id.slice(0, 6)
        }
      }))
      analyticsRef.current = aMap
      setAnalyticsMap(aMap)
      setRecordNames(recNameMap)

      // Build root record map: for each analytic, if there's a single root node
      // with children, store its ID. Used to fill missing analytics in Σ column keys.
      const rootMap: Record<string, string> = {}
      for (const [aid, nodes] of Object.entries(rMap)) {
        if (nodes.length === 1 && nodes[0].children.length > 0) {
          rootMap[aid] = nodes[0].record.id
        }
      }
      rootRecordByAidRef.current = rootMap

      const dbOrd = sa.map(b => b.analytic_id)
      dbOrdRef.current = dbOrd

      // Remember main analytic binding — used to extract indicator_id from
      // a coord_key when offering "promote to rule" snackbar after a per-cell
      // formula save.
      const mainBinding = sa.find(b => (b as any).is_main === 1 || (b as any).is_main === true)
      mainAnalyticIdRef.current = mainBinding?.analytic_id || null
      mainAnalyticIdxRef.current = mainBinding ? dbOrd.indexOf(mainBinding.analytic_id) : -1

      // View settings — read once per mount. Subsequent reloads reuse state.
      let curOrder = [...dbOrd]
      let effectivePinned = pinned
      if (!vsLoadedRef.current) {
        try {
          const vs: any = await api.getViewSettings(sheetId, currentUserId)
          if (vs?.order?.length) {
            const valid = new Set(dbOrd)
            curOrder = vs.order.filter((id: string) => valid.has(id))
            for (const id of dbOrd) if (!curOrder.includes(id)) curOrder.push(id)
          }
          if (vs?.pinned && Object.keys(vs.pinned).length > 0) {
            effectivePinned = vs.pinned
            setPinned(vs.pinned)
          }
          if (vs?.columnState && Array.isArray(vs.columnState) && vs.columnState.length > 0) {
            savedColumnStateRef.current = vs.columnState
          }
          if (vs?.colLevelToggles && typeof vs.colLevelToggles === 'object') {
            // Migrate old boolean format → new 3-state format
            const migrated: Record<number, ColLevelState> = {}
            for (const [k, v] of Object.entries(vs.colLevelToggles)) {
              if (typeof v === 'boolean') migrated[Number(k)] = v ? 'end' : 'hidden'
              else if (typeof v === 'string') migrated[Number(k)] = v as ColLevelState
            }
            setColLevelToggles(migrated)
          }
        } catch { /* ignore */ }
        vsLoadedRef.current = true
      }
      pinnedRef.current = effectivePinned

      const colAId = curOrder[0]
      colAIdRef.current = colAId
      const rowAIds = curOrder.slice(1).filter(id => !effectivePinned[id])

      const colTree = rMap[colAId] || []
      const colLeaves = getLeaves(colTree)
      colLeafIdsRef.current = colLeaves.map(l => l.record.id)
      colTreeRef.current = colTree

      // Detect labels for each non-leaf column level (Годы, Кварталы, …).
      const detectedLevels: { level: number; label: string }[] = []
      const walkForLabel = (nodes: RecordNode[], lvl: number) => {
        if (!nodes.length || !nodes[0].children.length) return
        const firstName = (nodes[0].data && nodes[0].data.name) || ''
        let label = `Уровень ${lvl + 1}`
        if (/^\d{4}$/.test(firstName)) label = 'Годы'
        else if (/квартал/i.test(firstName)) label = 'Кварталы'
        else if (/^(янв|фев|мар|апр|май|июн|июл|авг|сен|окт|ноя|дек)/i.test(firstName)) label = 'Месяцы'
        detectedLevels.push({ level: lvl, label })
        walkForLabel(nodes[0].children, lvl + 1)
      }
      walkForLabel(colTree, 0)
      setColLevelNames(detectedLevels)
      // Default all period levels to 'end' (show after leaves) unless user saved a preference.
      if (Object.keys(colLevelTogglesRef.current).length === 0 && detectedLevels.length > 0) {
        const defaults: Record<number, ColLevelState> = {}
        for (const { level } of detectedLevels) defaults[level] = 'end'
        setColLevelToggles(defaults)
        colLevelTogglesRef.current = defaults
      }

      // ── Column definitions (FLAT layout) ─────────────────────────────
      // Leaves first, then sum columns for each level based on toggle state.
      sumColLeavesRef.current = {}

      // Collect sum columns per level for 'start'/'end' placement
      const _collectSumCols = (nodes: RecordNode[], lvl: number, acc: Map<number, (ColDef | ColGroupDef)[]>) => {
        for (const n of nodes) {
          if (n.children.length === 0) continue
          const state = colLevelTogglesRef.current[lvl]
          if (state && state !== 'hidden') {
            const label = recordLabel(n)
            const leafIds = getLeaves([n]).map(l => l.record.id)
            if (!acc.has(lvl)) acc.set(lvl, [])
            acc.get(lvl)!.push(makeSumColDef(label, leafIds, n.record.id))
          }
          _collectSumCols(n.children, lvl + 1, acc)
        }
      }

      const buildColDefs = (nodes: RecordNode[], _lvl: number): (ColDef | ColGroupDef)[] => {
        // 1. Collect all leaf columns (flat)
        const leaves = getLeaves(nodes)
        const leafCols = leaves.map(l => makePeriodColDef(recordLabel(l), l.record.id))

        // 2. Collect sum columns grouped by level
        const sumByLevel = new Map<number, (ColDef | ColGroupDef)[]>()
        _collectSumCols(nodes, 0, sumByLevel)

        // 3. Build final order: 'start' sums → leaves → 'end' sums
        // Sum levels sorted from most granular (highest level) to least (lowest)
        const sortedLevels = Array.from(sumByLevel.keys()).sort((a, b) => b - a)
        const startCols: (ColDef | ColGroupDef)[] = []
        const endCols: (ColDef | ColGroupDef)[] = []
        for (const lvl of sortedLevels) {
          const state = colLevelTogglesRef.current[lvl]
          const cols = sumByLevel.get(lvl) || []
          if (state === 'start') startCols.push(...cols)
          else if (state === 'end') endCols.push(...cols)
        }

        return [...startCols, ...leafCols, ...endCols]
      }

      // ── Cells (full prefetch) ───────────────────────────────────────────
      cellMapRef.current = {}
      cellRuleRef.current = {}
      formulaMapRef.current = {}
      const cellData: CellData[] = await api.getCells(sheetId, currentUserId)
      for (const c of cellData) {
        cellMapRef.current[c.coord_key] = c.value ?? ''
        if (c.rule) cellRuleRef.current[c.coord_key] = c.rule as CellRule
        if (c.formula) formulaMapRef.current[c.coord_key] = c.formula
      }

      // ── Build rows (hierarchical, matching legacy PivotGrid) ────────────
      // We walk each row-analytic in order; within an analytic, we traverse
      // the parent_id tree; a node becomes a group row if it has children OR
      // if there's a sub-analytic below. Leaf rows (terminal, last analytic)
      // carry cell values.
      const rows: RowDatum[] = []
      const buildLevel = (
        ai: number,
        parentIds: Record<string, string>,
        pathSoFar: string[],
      ) => {
        if (ai >= rowAIds.length) return
        const aId = rowAIds[ai]
        const isLastAnalytic = ai === rowAIds.length - 1

        const walk = (nodes: RecordNode[], currentPath: string[]) => {
          for (const node of nodes) {
            const recIds = { ...parentIds, [aId]: node.record.id }
            const hasChildren = node.children.length > 0
            const isGroup = hasChildren || !isLastAnalytic
            const pathStep = `${aId}:${node.record.id}`
            const path = [...currentPath, pathStep]
            const row: RowDatum = {
              path,
              label: recordLabel(node),
              isLeaf: !isGroup,
              recordIds: recIds,
              unit: node.data?.unit,
              format: node.data?.format,
            }
            // Populate period cell values. For leaves we pull from cellMap;
            // for group rows the value is recomputed in a bottom-up pass
            // below (matches legacy sum_children default).
            for (const leaf of colLeaves) {
              const parts: string[] = []
              for (const a of dbOrd) {
                if (a === colAId) parts.push(leaf.record.id)
                else if (effectivePinned[a]) parts.push(effectivePinned[a])
                else if (recIds[a]) parts.push(recIds[a])
                else if (rootMap[a]) parts.push(rootMap[a])
              }
              if (parts.length >= 2) {
                const lookup = lookupCell(parts)
                row[`p_${leaf.record.id}`] = lookup.value
                row[`__coord_${leaf.record.id}`] = parts.join('|')
                row[`__parts_${leaf.record.id}`] = parts
                // 1. Full-key rule from backend (exact match)
                // 2. Truncated-key rule from lookupCell (when cell found via fallback)
                // 3. Default: sum_children for groups, manual for terminals
                const fullRule = cellRuleRef.current[parts.join('|')]
                row[`__rule_${leaf.record.id}`] = fullRule
                  ?? (lookup.key !== parts.join('|') && lookup.value !== '' ? lookup.rule : undefined)
                  ?? (isGroup ? 'sum_children' : 'manual')
              }
            }
            rows.push(row)
            if (hasChildren) walk(node.children, path)
            if (!hasChildren && !isLastAnalytic) {
              buildLevel(ai + 1, recIds, path)
            }
          }
        }
        walk(rMap[aId] || [], pathSoFar)
      }
      buildLevel(0, {}, [])

      // ── If ALL row analytics are pinned, buildLevel produced no rows.
      //    Render a single summary row so pinned values are still visible. ──
      if (rowAIds.length === 0 && Object.keys(effectivePinned).length > 0) {
        const labelParts = Object.entries(effectivePinned).map(([aId, rId]) =>
          `${aMap[aId]?.name || aId.slice(0, 6)}: ${recNameMap[rId] || '?'}`
        )
        const row: RowDatum = {
          path: ['__pinned_summary__'],
          label: labelParts.join(' / ') || '(зафиксировано)',
          isLeaf: true,
          recordIds: { ...effectivePinned },
        }
        for (const leaf of colLeaves) {
          const parts: string[] = []
          for (const a of dbOrd) {
            if (a === colAId) parts.push(leaf.record.id)
            else if (effectivePinned[a]) parts.push(effectivePinned[a])
          }
          if (parts.length >= 2) {
            const lookup = lookupCell(parts)
            row[`p_${leaf.record.id}`] = lookup.value
            row[`__coord_${leaf.record.id}`] = parts.join('|')
            row[`__parts_${leaf.record.id}`] = parts
            row[`__rule_${leaf.record.id}`] =
              cellRuleRef.current[parts.join('|')] ?? 'manual'
          }
        }
        rows.push(row)
      }

      // ── Bottom-up sum_children computation for group rows ───────────────
      const pathIdx: Record<string, number[]> = {}
      rows.forEach((r, i) => {
        const parentKey = r.path.slice(0, -1).join('|')
        ;(pathIdx[parentKey] ||= []).push(i)
      })
      const sumFor = (rowIdx: number, field: string): number | null => {
        const row = rows[rowIdx]
        const leafId = field.slice(2)
        const rule = row[`__rule_${leafId}`]
        if (rule === 'empty') return null
        if (row.isLeaf) {
          const v = row[field]
          if (v == null || v === '') return null
          const n = parseFloat(String(v))
          return Number.isNaN(n) ? null : n
        }
        const kids = pathIdx[row.path.join('|')] || []
        let s = 0, has = false
        for (const ci of kids) {
          const sv = sumFor(ci, field)
          if (sv != null) { s += sv; has = true }
        }
        return has ? s : null
      }
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].isLeaf) continue
        for (const leaf of colLeaves) {
          const field = `p_${leaf.record.id}`
          const rule = rows[i][`__rule_${leaf.record.id}`]
          if (rule === 'empty') { rows[i][field] = ''; continue }
          // Only compute if current value is empty AND rule is sum_children
          // (explicit formula/manual rules keep their stored value).
          if ((rows[i][field] === '' || rows[i][field] == null) && rule === 'sum_children') {
            const s = sumFor(i, field)
            if (s != null) rows[i][field] = String(s)
          }
        }
      }

      // ── Column defs ─────────────────────────────────────────────────────
      const defs: (ColDef | ColGroupDef)[] = buildColDefs(colTree, 0)
      setColumnDefs(defs)
      setRowData(rows)
      setSheetName(`${rowAIds.length} аналитик × ${colLeaves.length} периодов`)
      setLoading(false)
    } catch (e: any) {
      setError(e?.message || String(e))
      setLoading(false)
    }
  }, [sheetId, currentUserId, lookupCell, pinned])

  useEffect(() => { load() }, [load])

  // In formulas mode, resolve indicator-rule formulas for all known coord keys
  // (so HEAD cells driven by consolidation/scoped rules show their formula +
  // source badge in the grid instead of a generic "Σ сумма" label).
  useEffect(() => {
    if (!sheetId) return
    const keys = Object.keys(cellMapRef.current)
    if (keys.length === 0) return
    let cancelled = false
    api.getResolvedFormulas(sheetId, keys).then(res => {
      if (cancelled) return
      const m: Record<string, { formula: string; source: string; kind: string }> = {}
      for (const r of res) if (r.formula) m[r.coord_key] = { formula: r.formula, source: r.source, kind: r.kind }
      resolvedFormulaMapRef.current = m
      gridApiRef.current?.refreshCells({ force: true })
    }).catch(() => {})
    return () => { cancelled = true }
  }, [sheetId, mode, rowData])

  // Rebuild columnDefs (with / without sum columns) when level toggles change.
  // Uses the cached colTree so we don't re-fetch cells.
  const togglesSig = JSON.stringify(colLevelToggles)
  useEffect(() => {
    const colTree = colTreeRef.current
    if (!colTree.length) return
    // Reset sum-col leaf map — toggle rebuild re-registers them below.
    sumColLeavesRef.current = {}
    const rebuild = (nodes: RecordNode[], _lvl: number): (ColDef | ColGroupDef)[] => {
      // Same flat layout as buildColDefs
      sumColLeavesRef.current = {}
      const leaves = getLeaves(nodes)
      const leafCols = leaves.map(l => makePeriodColDef(recordLabel(l), l.record.id))

      const sumByLevel = new Map<number, (ColDef | ColGroupDef)[]>()
      const collectSums = (ns: RecordNode[], lvl: number) => {
        for (const n of ns) {
          if (n.children.length === 0) continue
          const state = colLevelToggles[lvl]
          if (state && state !== 'hidden') {
            const label = recordLabel(n)
            const leafIds = getLeaves([n]).map(l => l.record.id)
            if (!sumByLevel.has(lvl)) sumByLevel.set(lvl, [])
            sumByLevel.get(lvl)!.push(makeSumColDef(label, leafIds, n.record.id))
          }
          collectSums(n.children, lvl + 1)
        }
      }
      collectSums(nodes, 0)

      const sortedLevels = Array.from(sumByLevel.keys()).sort((a, b) => b - a)
      const startCols: (ColDef | ColGroupDef)[] = []
      const endCols: (ColDef | ColGroupDef)[] = []
      for (const lvl of sortedLevels) {
        const state = colLevelToggles[lvl]
        const cols = sumByLevel.get(lvl) || []
        if (state === 'start') startCols.push(...cols)
        else if (state === 'end') endCols.push(...cols)
      }

      return [...startCols, ...leafCols, ...endCols]
    }
    setColumnDefs(prev => {
      // Preserve the auto-group column (always first, pinned left) + any
      // currently-injected row-analytic columns (none in this grid). We
      // only regenerate columns that come from the period tree.
      const first = prev.find(c => (c as ColGroupDef).children == null && (c as ColDef).field === undefined)
      const newCols = rebuild(colTree, 0)
      return first ? [first, ...newCols] : newCols
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [togglesSig])

  // Persist view settings (pinned analytics + column state: widths, order,
  // pinned cols, hidden cols). Debounced, only after first load.
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  useEffect(() => {
    if (!vsLoadedRef.current) return
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
    saveTimerRef.current = setTimeout(() => {
      api.saveViewSettings(sheetId, {
        pinned,
        columnState: savedColumnStateRef.current,
        colLevelToggles,
        _user_id: currentUserId,
      }).catch(() => { /* ignore */ })
    }, 500)
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
    }
  }, [pinned, columnStateVersion, sheetId, colLevelToggles, currentUserId])

  const handlePin = useCallback((analyticId: string, recordId: string) => {
    // Don't pin the column analytic — it's the grid's column axis.
    if (analyticId === colAIdRef.current) return
    setPinned(prev => ({ ...prev, [analyticId]: recordId }))
  }, [])

  const handleUnpin = useCallback((analyticId: string) => {
    setPinned(prev => {
      const next = { ...prev }
      delete next[analyticId]
      return next
    })
  }, [])

  // Fetch fresh cells from server, update leaves silently, bottom-up
  // recompute parents, and flash parents whose value changed. Reused by
  // the model-wide recalc effect AND by onCellValueChanged (so formula
  // cells recomputed server-side after a local edit also show the flash).
  const refreshAndFlashParents = useCallback(async () => {
    const grid = gridApiRef.current
    if (!grid) return
    try {
      const fresh: CellData[] = await api.getCells(sheetId, currentUserId)
      const freshMap: Record<string, string> = {}
      const freshRule: Record<string, CellRule> = {}
      const freshFormula: Record<string, string> = {}
      for (const c of fresh) {
        freshMap[c.coord_key] = c.value ?? ''
        if (c.rule) freshRule[c.coord_key] = c.rule as CellRule
        if (c.formula) freshFormula[c.coord_key] = c.formula
      }
      formulaMapRef.current = freshFormula
      // Pass 1: update every LEAF row's period values silently.
      // Use direct data mutation + refreshCells (NOT setDataValue which
      // would re-trigger onCellValueChanged and cause save loops).
      const changedLeafNodes: any[] = []
      const changedLeafCols = new Set<string>()
      grid.forEachNode(node => {
        if (!node.data || !node.data.isLeaf) return
        let rowChanged = false
        for (const leafId of colLeafIdsRef.current) {
          const coordKey: string | undefined = node.data[`__coord_${leafId}`]
          if (!coordKey) continue
          const newVal = freshMap[coordKey] ?? ''
          const oldVal = (node.data[`p_${leafId}`] ?? '') as string | number
          const newRule = (freshRule[coordKey] as CellRule) || 'manual'
          if (newRule !== node.data[`__rule_${leafId}`]) {
            node.data[`__rule_${leafId}`] = newRule
          }
          if (String(newVal) !== String(oldVal)) {
            node.data[`p_${leafId}`] = newVal
            rowChanged = true
            changedLeafCols.add(`p_${leafId}`)
          }
        }
        if (rowChanged) changedLeafNodes.push(node)
      })
      if (changedLeafNodes.length > 0) {
        grid.refreshCells({ rowNodes: changedLeafNodes, columns: Array.from(changedLeafCols), force: true })
      }
      // Pass 2: recompute parent (non-leaf) rows bottom-up; flash changes.
      const allNodes: any[] = []
      grid.forEachNode(n => { if (n.data) allNodes.push(n) })
      const kidsByPath: Record<string, any[]> = {}
      for (const n of allNodes) {
        const pk = n.data.path.slice(0, -1).join('|')
        ;(kidsByPath[pk] ||= []).push(n)
      }
      const parents = allNodes
        .filter(n => !n.data.isLeaf)
        .sort((a, b) => b.data.path.length - a.data.path.length)
      const flashRowNodes: any[] = []
      const flashColSet = new Set<string>()
      for (const pNode of parents) {
        const kids = kidsByPath[pNode.data.path.join('|')] || []
        let rowChanged = false
        for (const leafId of colLeafIdsRef.current) {
          const field = `p_${leafId}`
          const coordKey: string | undefined = pNode.data[`__coord_${leafId}`]
          if (coordKey && freshRule[coordKey]) {
            pNode.data[`__rule_${leafId}`] = freshRule[coordKey]
          }
          const rule: CellRule = pNode.data[`__rule_${leafId}`] || 'sum_children'
          let newVal: string
          if (rule === 'empty') {
            newVal = ''
          } else if (rule === 'formula' || rule === 'manual') {
            newVal = coordKey ? (freshMap[coordKey] ?? '') : ''
          } else {
            let s = 0, has = false
            for (const k of kids) {
              const v = k.data[field]
              if (v == null || v === '') continue
              const n = parseFloat(String(v))
              if (!Number.isNaN(n)) { s += n; has = true }
            }
            newVal = has ? String(s) : ''
          }
          const oldVal = (pNode.data[field] ?? '') as string | number
          if (String(newVal) !== String(oldVal)) {
            pNode.data[field] = newVal
            rowChanged = true
            flashColSet.add(field)
          }
        }
        if (rowChanged) flashRowNodes.push(pNode)
      }
      if (flashRowNodes.length > 0 && flashColSet.size > 0) {
        grid.refreshCells({ rowNodes: flashRowNodes, columns: Array.from(flashColSet), force: true })
        grid.flashCells({
          rowNodes: flashRowNodes,
          columns: Array.from(flashColSet),
          flashDuration: 1500,
          fadeDuration: 600,
        })
      }
      cellMapRef.current = freshMap
      cellRuleRef.current = freshRule
      // Refresh Σ (sum) columns so their valueGetters pick up server-computed
      // consolidation values from the updated cellMapRef.
      const sumColIds = Object.keys(sumColLeavesRef.current)
      if (sumColIds.length > 0) {
        grid.refreshCells({ columns: sumColIds, force: true })
      }
    } catch (e) { console.error('[flash] error', e) }
  }, [sheetId, currentUserId])

  // When a model-wide recalc ends (calcProgress truthy → null), refetch and
  // flash derived parent cells that changed.
  const prevCalcProgressRef = useRef(calcProgress)
  useEffect(() => {
    const wasRunning = !!prevCalcProgressRef.current
    const nowIdle = !calcProgress
    prevCalcProgressRef.current = calcProgress
    if (!wasRunning || !nowIdle) return
    refreshAndFlashParents()
  }, [calcProgress, refreshAndFlashParents])

  // Native clipboard paste — reads TSV from the paste event (works in all
  // browsers without permission prompt because the paste event is
  // user-initiated). Writes into manual cells starting at focused cell.
  const gridContainerRef = useRef<HTMLDivElement | null>(null)
  useEffect(() => {
    const el = gridContainerRef.current
    if (!el) return
    const h = (e: ClipboardEvent) => {
      const grid = gridApiRef.current
      if (!grid) return
      // If user is editing a cell, let the browser handle paste into the input.
      if (grid.getEditingCells()?.length) return
      const text = e.clipboardData?.getData('text/plain') || ''
      if (!text) return
      e.preventDefault()
      e.stopPropagation()
      const matrix = text
        .replace(/\r\n/g, '\n')
        .split('\n')
        .filter((l, i, arr) => !(i === arr.length - 1 && l === ''))
        .map(line => line.split('\t'))
      if (matrix.length === 0) return
      const focused = grid.getFocusedCell()
      if (!focused) return
      const startRow = focused.rowIndex
      const startColId = focused.column.getColId()
      // Gather all displayed columns to walk rightwards from startColId.
      const allCols = grid.getAllDisplayedColumns()
      const startColIdx = allCols.findIndex(c => c.getColId() === startColId)
      if (startColIdx < 0) return
      const toSave: { coord_key: string; value: string; data_type: string; rule: 'manual'; user_id?: string }[] = []
      for (let dr = 0; dr < matrix.length; dr++) {
        const row = grid.getDisplayedRowAtIndex(startRow + dr)
        if (!row?.data || !row.data.isLeaf) continue
        const rowCells = matrix[dr]
        for (let dc = 0; dc < rowCells.length; dc++) {
          const col = allCols[startColIdx + dc]
          if (!col) break
          const colId = col.getColId()
          if (!colId.startsWith('p_')) continue
          const leafId = colId.slice(2)
          const rule: CellRule = row.data[`__rule_${leafId}`] || 'manual'
          if (rule !== 'manual') continue
          // Normalize pasted value: strip spaces/nbsp, convert comma decimals.
          let v = String(rowCells[dc] ?? '').trim().replace(/\u00a0/g, '').replace(/\s+/g, '')
          if (v !== '') {
            const n = Number(v.replace(',', '.'))
            if (!Number.isNaN(n)) v = String(n)
            else continue // non-numeric cell — skip, don't pollute data
          }
          row.setDataValue(colId, v)
          const coordKey: string | undefined = row.data[`__coord_${leafId}`]
          if (coordKey) {
            toSave.push({
              coord_key: coordKey,
              value: v,
              data_type: 'number',
              rule: 'manual',
              user_id: currentUserId,
            })
          }
        }
      }
      if (toSave.length > 0) {
        setRecalcRunning(true)
        // Batch save — one round-trip for the whole paste.
        api.saveCells(sheetId, toSave)
          .then(() => {
            for (const c of toSave) cellMapRef.current[c.coord_key] = c.value
          })
          .catch((err: any) => setError(`Не удалось сохранить вставку: ${err?.message || err}`))
          .finally(() => setRecalcRunning(false))
      }
    }
    el.addEventListener('paste', h)
    return () => { el.removeEventListener('paste', h) }
  }, [sheetId, currentUserId])

  // ── Column def factory (kept outside to stay stable) ──────────────────
  function makePeriodColDef(label: string, periodRecId: string): ColDef {
    const field = `p_${periodRecId}`
    // Rough width from header text length so names fit by default.
    // Generous by design: AG Grid's autoSizeColumns-to-header measurement
    // is fragile after column-state restore, so we err on the wide side so
    // labels like "Декабрь 2026" don't get clipped. 12 px/char + 44 px of
    // padding (cell padding + sort/menu icon) ≈ fits "Декабрь 2026" (12 ch
    // → 188 px) with margin.
    const chars = (label || '').length
    const autoWidth = Math.max(120, Math.min(320, chars * 12 + 44))
    return {
      headerName: label,
      field,
      headerClass: 'ag-center-header',
      width: autoWidth,
      minWidth: 90,
      wrapText: true,
      autoHeight: true,
      // NB: do NOT set enableCellChangeFlash here — we only want to flash
      // DERIVED cells (sums/formulas that recomputed), not the manual cell
      // the user just typed into. Flash is triggered explicitly via
      // api.flashCells() in the recalc-diff effect and in the local
      // parent-sum update path.
      editable: (p: any) => {
        if (mode === 'formulas') return false // read-only in formula view
        if (!p.data || !p.data.isLeaf) return false // group rows: read-only
        const rule: CellRule = p.data[`__rule_${periodRecId}`] || 'manual'
        // Только `manual` — редактируемо. Клетки с формулой редактируются
        // через кнопку ⋮ → FormulaEditor (а не путём ввода значения), чтобы
        // случайный ввод не затирал формулу. `empty` / `sum_children` —
        // read-only.
        return rule === 'manual'
      },
      valueFormatter: (p: any) => {
        const rule: CellRule = p.data?.[`__rule_${periodRecId}`] || 'manual'
        if (mode === 'formulas') {
          // Formula mode: show resolved indicator-rule formula (single source of truth).
          const coordKey: string | undefined = p.data?.[`__coord_${periodRecId}`]
          const resolved = coordKey ? resolvedFormulaMapRef.current[coordKey] : undefined
          if (resolved?.formula) {
            return resolved.formula
          }
          if (rule === 'formula') return 'ƒ формула'
          if (rule === 'sum_children') return 'ƒ SUM'
          if (rule === 'empty') return '∅'
          if (!p.data?.isLeaf) return 'ƒ SUM'
          return '✎ ввод'
        }
        if (rule === 'empty') return ''
        const v = p.value
        if (v == null || v === '') return ''
        const num = Number(v)
        if (!Number.isNaN(num)) {
          const fmt = p.data?.format
          if (fmt === 'percent') {
            return (num * 100).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%'
          }
          return num.toLocaleString('ru-RU', { maximumFractionDigits: 2 })
        }
        return String(v)
      },
      valueParser: (p: any) => {
        // Lenient numeric parser: strip everything except digits and the
        // first decimal separator. Example: "11у12.а12" → "1112.12".
        // Also: comma → dot, minus sign allowed only at position 0.
        const raw = String(p.newValue ?? '').trim()
        if (raw === '') return ''
        // Normalise decimal comma to dot before filtering.
        const normalised = raw.replace(',', '.')
        let out = ''
        let seenDot = false
        for (let i = 0; i < normalised.length; i++) {
          const ch = normalised[i]
          if (ch >= '0' && ch <= '9') {
            out += ch
          } else if (ch === '.' && !seenDot) {
            out += '.'
            seenDot = true
          } else if (ch === '-' && out === '') {
            out += '-'
          }
          // everything else gets dropped
        }
        // Cleanup: bare '-' / '.' / '-.' → empty
        if (out === '' || out === '-' || out === '.' || out === '-.') return ''
        let num = Number(out)
        if (Number.isNaN(num)) return p.oldValue
        // For percent cells: user types "15" meaning 15%, store as 0.15
        const fmt = p.data?.format
        if (fmt === 'percent') {
          num = num / 100
          return String(num)
        }
        return out
      },
      cellStyle: (p: any) => {
        const isLeaf = !!p.data?.isLeaf
        const rule: CellRule = p.data?.[`__rule_${periodRecId}`] || 'manual'
        const s: any = { textAlign: mode === 'formulas' ? 'left' : 'right' }
        if (mode === 'formulas') {
          // Settings/formula view: distinct pastel palette + left align so
          // rule labels and formula text read naturally.
          s.fontSize = 11
          s.background = '#fafbfc'
          s.whiteSpace = 'normal'
          s.lineHeight = '1.4'
          s.overflow = 'visible'
          s.paddingTop = '4px'
          if (rule === 'formula') s.color = '#1565c0'
          else if (rule === 'sum_children' || !isLeaf) s.color = '#2e7d32'
          else if (rule === 'empty') s.color = '#bbb'
          else s.color = '#666'
          return s
        }
        if (rule === 'empty') {
          s.background = '#f5f5f5'
          s.color = '#bbb'
          return s
        }
        if (!isLeaf || rule === 'sum_children') {
          s.background = '#fafafa'
          s.color = '#2e7d32'
          s.fontWeight = 500
        } else if (rule === 'formula') {
          // White background for all formula cells (including zero values);
          // blue text is enough to distinguish them from manual input.
          s.background = '#fff'
          s.color = '#1565c0'
        } else {
          s.background = '#fdf8e8' // manual input tint (legacy parity)
        }
        const num = Number(p.value)
        if (!Number.isNaN(num) && num < 0) s.color = '#d32f2f'
        return s
      },
      tooltipValueGetter: (p: any) => {
        const coordKey: string | undefined = p.data?.[`__coord_${periodRecId}`]
        if (!coordKey) return null
        const rule: CellRule = p.data?.[`__rule_${periodRecId}`] || 'manual'
        const isLeaf = !!p.data?.isLeaf
        // Resolved indicator rule formula (single source of truth)
        const resolved = resolvedFormulaMapRef.current[coordKey]
        if (resolved?.formula) return `ƒ ${resolved.formula}`
        // Consolidating cells (group indicator or period parent)
        if (!isLeaf || rule === 'sum_children') return 'ƒ SUM'
        return null
      },
      tooltipComponent: FormulaTooltip,
      // Hover ⋮ button → opens FormulaEditor for this cell.
      // Only on leaf rows (group rows don't have an editable formula).
      cellRenderer: (p: any) => {
        const isLeaf = !!p.data?.isLeaf
        const coordKey: string | undefined = p.data?.[`__coord_${periodRecId}`]
        const formatted = p.valueFormatted != null && p.valueFormatted !== ''
          ? p.valueFormatted
          : (p.value == null ? '' : String(p.value))
        if (!coordKey || mode !== 'formulas') {
          return <span>{formatted}</span>
        }
        return (
          <span
            className="cell-with-menu"
            style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', width: '100%', gap: 4 }}
          >
            <span style={{ flex: 1, wordBreak: 'break-word', whiteSpace: 'normal' }}>
              {formatted}
            </span>
            <button
              type="button"
              className="cell-menu-btn"
              title="Формула клетки"
              onClick={ev => {
                ev.stopPropagation()
                setFormulaEditorKey(coordKey)
                setFormulaEditorOpen(true)
              }}
              style={{
                border: 'none', background: 'transparent', cursor: 'pointer',
                padding: 0, lineHeight: 1, color: '#888', opacity: 0,
                transition: 'opacity 0.15s', flexShrink: 0,
              }}
            >
              <MoreVertOutlined sx={{ fontSize: 14 }} />
            </button>
          </span>
        )
      },
    }
  }

  // Level-aggregate column (e.g. year total across quarters). Non-editable,
  // summed from the group's leaf period fields for the current row.
  function makeSumColDef(groupLabel: string, leafIds: string[], groupRecordId?: string): ColDef {
    /** Build coord_key for Σ column, filling missing analytics from rootRecordByAidRef. */
    const buildSumKey = (recordIds: Record<string, string> | undefined): string | null => {
      if (!groupRecordId || !recordIds) return null
      const dbOrd = dbOrdRef.current
      const colAId = colAIdRef.current
      const roots = rootRecordByAidRef.current
      const parts: string[] = []
      for (const a of dbOrd) {
        if (a === colAId) parts.push(groupRecordId)
        else {
          const rid = recordIds[a] || roots[a]
          if (rid) parts.push(rid)
        }
      }
      return parts.length === dbOrd.length ? parts.join('|') : null
    }
    const label = groupLabel
    // Stable colId so we can target flashing — suffixed with leaf-id joins so
    // two different Σ columns covering different leaf sets stay distinct.
    const colId = `sum__${leafIds.join('_').slice(0, 80)}__${leafIds.length}`
    sumColLeavesRef.current[colId] = leafIds
    return {
      colId,
      headerName: label,
      headerClass: 'ag-center-header',
      width: Math.max(100, label.length * 9 + 24),
      minWidth: 80,
      editable: false,
      valueGetter: (p: any) => {
        if (!p.data) return ''
        // Server-computed consolidation value — single source of truth.
        // The formula engine computes these from indicator_formula_rules,
        // including for group indicator rows (sum of children).
        {
          const key = buildSumKey(p.data.recordIds)
          if (key) {
            const serverVal = cellMapRef.current[key]
            if (serverVal != null && serverVal !== '') {
              const n = parseFloat(String(serverVal))
              if (!Number.isNaN(n)) return n
            }
          }
        }
        // Fallback: client-side SUM when no server cell exists yet
        // (e.g. before first recalc).
        let s = 0, has = false
        for (const id of leafIds) {
          const v = p.data[`p_${id}`]
          if (v == null || v === '') continue
          const n = parseFloat(String(v))
          if (!Number.isNaN(n)) { s += n; has = true }
        }
        return has ? s : ''
      },
      valueFormatter: (p: any) => {
        if (mode === 'formulas') {
          const key = buildSumKey(p.data?.recordIds)
          if (key) {
            const resolved = resolvedFormulaMapRef.current[key]
            if (resolved?.formula) return resolved.formula
          }
          return 'ƒ SUM'
        }
        const v = p.value
        if (v === '' || v == null) return ''
        const num = Number(v)
        if (Number.isNaN(num)) return String(v)
        return num.toLocaleString('ru-RU', { maximumFractionDigits: 2 })
      },
      cellStyle: (): any => ({
        textAlign: mode === 'formulas' ? 'left' : 'right',
        background: '#eef6ff',
        color: '#0d47a1',
        fontWeight: mode === 'formulas' ? 400 : 600,
        ...(mode === 'formulas' ? { fontSize: 11, fontFamily: 'monospace', whiteSpace: 'normal', lineHeight: '1.4', overflow: 'visible' } : {}),
      }),
      tooltipValueGetter: (p: any) => {
        const key = buildSumKey(p.data?.recordIds)
        if (!key) return null
        const resolved = resolvedFormulaMapRef.current[key]
        if (resolved?.formula) return `ƒ ${resolved.formula}`
        return 'ƒ SUM'
      },
      tooltipComponent: FormulaTooltip,
    }
  }

  const getDataPath = useCallback((data: RowDatum) => data.path, [])

  const autoGroupColumnDef = useMemo<ColDef>(() => ({
    headerName: 'Аналитика',
    minWidth: 280,
    pinned: 'left',
    wrapText: true,
    autoHeight: true,
    cellRendererParams: {
      suppressCount: true,
      innerRenderer: (p: any) => {
        const label = p.data?.label ?? ''
        const unit = p.data?.unit ? ` (${p.data.unit})` : ''
        const isGroup = !p.data?.isLeaf
        const path: string[] = p.data?.path || []
        const lastStep = path[path.length - 1] || ''
        const idx = lastStep.indexOf(':')
        const aId = idx > 0 ? lastStep.slice(0, idx) : ''
        const recId = idx > 0 ? lastStep.slice(idx + 1) : ''
        const canDrag = !!aId && !!recId
        return (
          <span
            draggable={canDrag}
            onDragStart={(e) => {
              if (!canDrag) return
              e.dataTransfer.setData('text/plain', JSON.stringify({ analyticId: aId, recordId: recId }))
              e.dataTransfer.effectAllowed = 'copy'
            }}
            style={{
              fontWeight: isGroup ? 600 : 400,
              cursor: canDrag ? 'grab' : 'default',
              userSelect: 'none',
            }}
            title={canDrag ? 'Перетащите на панель сверху, чтобы зафиксировать аналитику' : undefined}
          >
            {label}{unit}
          </span>
        )
      },
    },
    cellStyle: (p: any) => {
      const isGroup = !p.data?.isLeaf
      return {
        background: isGroup ? '#fafafa' : '#fff',
        fontWeight: isGroup ? 600 : 400,
      }
    },
  }), [])

  const onGridReady = useCallback((e: GridReadyEvent) => {
    gridApiRef.current = e.api
    // Expose for E2E tests. Harmless in production — just a handle.
    ;(window as unknown as { __pebbleGridApi?: GridApi }).__pebbleGridApi = e.api
  }, [])

  // Apply saved column state every time columnDefs are (re)built. If no
  // saved state exists, auto-size columns to fit headers.
  useEffect(() => {
    if (columnDefs.length === 0) return
    const grid = gridApiRef.current
    if (!grid) return
    // Defer to next frame so AG Grid has registered the new columns.
    const raf = requestAnimationFrame(() => {
      const g = gridApiRef.current
      if (!g) return
      if (savedColumnStateRef.current && savedColumnStateRef.current.length > 0) {
        g.applyColumnState({
          state: savedColumnStateRef.current,
          applyOrder: true,
        })
      }
      // Note: we DON'T call autoSizeColumns here. AG Grid's header-measurement
      // has been unreliable for long localised labels (e.g. "Декабрь 2026"
      // getting clipped to ~138px). We instead rely on the generous default
      // `width` in makePeriodColDef so headers always fit out of the box;
      // users can shrink via drag, which persists via column-state.
    })
    return () => cancelAnimationFrame(raf)
  }, [columnDefs])

  // Capture current column state → ref + bump version to trigger save effect.
  const captureColumnState = useCallback(() => {
    const grid = gridApiRef.current
    if (!grid) return
    savedColumnStateRef.current = grid.getColumnState()
    setColumnStateVersion(v => v + 1)
  }, [])

  // Recompute selection stats (count / sum / avg / min / max) for the current
  // cell ranges and publish to the external store that feeds the status panel.
  const updateSelectionStats = useCallback(() => {
    const api = gridApiRef.current
    if (!api) {
      selectionStore.set({ count: 0, numCount: 0, sum: 0, avg: 0, min: null, max: null })
      return
    }
    const ranges = api.getCellRanges() || []
    let count = 0, numCount = 0, sum = 0
    let min = Infinity, max = -Infinity
    for (const r of ranges) {
      const cols = r.columns || []
      const si = r.startRow?.rowIndex ?? 0
      const ei = r.endRow?.rowIndex ?? si
      const rMin = Math.min(si, ei), rMax = Math.max(si, ei)
      for (let ri = rMin; ri <= rMax; ri++) {
        const node = api.getDisplayedRowAtIndex(ri)
        if (!node?.data) continue
        for (const col of cols) {
          count++
          const v = (node.data as any)[col.getColId()]
          if (v == null || v === '') continue
          const n = typeof v === 'number' ? v : parseFloat(String(v))
          if (!Number.isNaN(n) && isFinite(n)) {
            numCount++; sum += n
            if (n < min) min = n
            if (n > max) max = n
          }
        }
      }
    }
    selectionStore.set({
      count, numCount, sum,
      avg: numCount > 0 ? sum / numCount : 0,
      min: numCount > 0 ? min : null,
      max: numCount > 0 ? max : null,
    })
  }, [])

  // ── Ctrl/Cmd+D / Ctrl/Cmd+R: Excel-style fill ─────────────────────────
  // D: copy first row of selection DOWN across the range
  // R: copy first column of selection RIGHT across the range
  // Only writes to manual cells (formula/sum are left untouched).
  const onCellKeyDown = useCallback((e: CellKeyDownEvent) => {
    const kev = e.event as KeyboardEvent | null
    if (!kev) return
    const key = kev.key.toLowerCase()

    // Printable single char on a focused (non-editing) cell → start editing
    // fresh with that char (Excel behavior: replace, don't append).
    if (
      !kev.ctrlKey && !kev.metaKey && !kev.altKey &&
      kev.key.length === 1 &&
      !(e.api.getEditingCells()?.length)
    ) {
      const focused = e.api.getFocusedCell()
      if (focused?.column) {
        const colId = focused.column.getColId()
        const node = e.api.getDisplayedRowAtIndex(focused.rowIndex)
        if (node?.data?.isLeaf && colId.startsWith('p_')) {
          const rule: CellRule = node.data[`__rule_${colId.slice(2)}`] || 'manual'
          // Allow keyboard-to-edit for manual AND formula cells (Excel parity:
          // typing on a formula cell replaces it with manual value).
          if (rule === 'manual' || rule === 'formula') {
            kev.preventDefault()
            kev.stopPropagation()
            e.api.startEditingCell({
              rowIndex: focused.rowIndex,
              colKey: colId,
              key: kev.key,
            })
            return
          }
        }
      }
    }

    // Alt+ArrowLeft / Alt+ArrowRight → collapse/expand focused group row.
    if (kev.altKey && (key === 'arrowleft' || key === 'arrowright')) {
      const focused = e.api.getFocusedCell()
      if (!focused) return
      const node = e.api.getDisplayedRowAtIndex(focused.rowIndex)
      if (!node || !node.group) return
      kev.preventDefault()
      kev.stopPropagation()
      node.setExpanded(key === 'arrowright')
      return
    }

    // Cmd/Ctrl+ArrowLeft — collapse the parent of the current row and move
    // focus up to that parent. From a leaf row, folds the enclosing group.
    if ((kev.metaKey || kev.ctrlKey) && key === 'arrowleft') {
      const focused = e.api.getFocusedCell()
      if (!focused) return
      const cur = e.api.getDisplayedRowAtIndex(focused.rowIndex)
      const curPath: string[] | undefined = cur?.data?.path
      if (!cur || !curPath || curPath.length < 2) return
      // Walk upward through displayed rows to find the nearest ancestor
      // (a row whose path is a strict prefix of the current row's path).
      let parentIdx = -1
      for (let i = focused.rowIndex - 1; i >= 0; i--) {
        const n = e.api.getDisplayedRowAtIndex(i)
        const p: string[] | undefined = n?.data?.path
        if (!p) continue
        if (p.length < curPath.length && curPath.slice(0, p.length).join('|') === p.join('|')) {
          parentIdx = i
          break
        }
      }
      if (parentIdx < 0) return
      const parent = e.api.getDisplayedRowAtIndex(parentIdx)
      if (!parent) return
      kev.preventDefault()
      kev.stopPropagation()
      if (parent.group) parent.setExpanded(false)
      e.api.setFocusedCell(parentIdx, focused.column.getColId())
      e.api.ensureIndexVisible(parentIdx)
      return
    }

    const apiRef = e.api
    const isPeriodCol = (id: string) => id.startsWith('p_')
    const ruleOf = (data: any, colId: string) =>
      (data?.[`__rule_${colId.slice(2)}`] as CellRule) || 'manual'

    // Delete / Backspace → clear all manual cells in the current range.
    if (key === 'delete' || key === 'backspace') {
      const ranges = apiRef.getCellRanges()
      if (!ranges || ranges.length === 0) return
      kev.preventDefault()
      kev.stopPropagation()
      for (const range of ranges) {
        const cols = range.columns
        if (!cols || cols.length === 0) continue
        const si = range.startRow?.rowIndex ?? 0
        const ei = range.endRow?.rowIndex ?? si
        const rMin = Math.min(si, ei), rMax = Math.max(si, ei)
        for (let r = rMin; r <= rMax; r++) {
          const node = apiRef.getDisplayedRowAtIndex(r)
          if (!node?.data || !node.data.isLeaf) continue
          for (const col of cols) {
            const colId = col.getColId()
            if (!isPeriodCol(colId)) continue
            if (ruleOf(node.data, colId) !== 'manual') continue
            if (node.data[colId] === '' || node.data[colId] == null) continue
            node.setDataValue(colId, '')
          }
        }
      }
      return
    }

    const mod = kev.ctrlKey || kev.metaKey
    if (!mod) return

    // Cmd/Ctrl+A → select all period columns × all leaf rows.
    if (key === 'a') {
      kev.preventDefault()
      kev.stopPropagation()
      const total = apiRef.getDisplayedRowCount()
      if (total === 0) return
      // Gather all visible period column IDs.
      const allCols = apiRef.getAllDisplayedColumns()
      const periodColIds = allCols
        .map(c => c.getColId())
        .filter(isPeriodCol)
      if (periodColIds.length === 0) return
      apiRef.clearRangeSelection()
      apiRef.addCellRange({
        rowStartIndex: 0,
        rowEndIndex: total - 1,
        columns: periodColIds,
      })
      return
    }

    if (key !== 'd' && key !== 'r') return
    kev.preventDefault()
    kev.stopPropagation()
    const ranges = apiRef.getCellRanges()
    if (!ranges || ranges.length === 0) return
    const range = ranges[0]
    const cols = range.columns
    if (!cols || cols.length === 0) return
    const si = range.startRow?.rowIndex ?? 0
    const ei = range.endRow?.rowIndex ?? si
    const rMin = Math.min(si, ei), rMax = Math.max(si, ei)
    if (key === 'd') {
      const source = apiRef.getDisplayedRowAtIndex(rMin)
      if (!source?.data) return
      for (const col of cols) {
        const colId = col.getColId()
        if (!isPeriodCol(colId)) continue
        const val = source.data[colId]
        for (let r = rMin + 1; r <= rMax; r++) {
          const node = apiRef.getDisplayedRowAtIndex(r)
          if (!node?.data || !node.data.isLeaf) continue
          if (ruleOf(node.data, colId) !== 'manual') continue
          node.setDataValue(colId, val)
        }
      }
    } else {
      for (let r = rMin; r <= rMax; r++) {
        const node = apiRef.getDisplayedRowAtIndex(r)
        if (!node?.data || !node.data.isLeaf) continue
        const firstCol = cols[0].getColId()
        if (!isPeriodCol(firstCol)) continue
        const val = node.data[firstCol]
        for (let c = 1; c < cols.length; c++) {
          const colId = cols[c].getColId()
          if (!isPeriodCol(colId)) continue
          if (ruleOf(node.data, colId) !== 'manual') continue
          node.setDataValue(colId, val)
        }
      }
    }
  }, [])

  // Recalculate group-row sums in-place after a leaf cell changes.
  // Bottom-up walk from the parent up, updating all ancestor rows' cell
  // for the column that just changed. Non-blocking, local-only — server
  // recalc still runs in background for formula cells.
  const recomputeParentsForField = useCallback((leafPath: string[], field: string) => {
    const api = gridApiRef.current
    if (!api) return
    // Build fresh node→children index from currently displayed rows.
    const rows: RowDatum[] = []
    api.forEachNode(n => { if (n.data) rows.push(n.data as RowDatum) })
    const byPath: Record<string, RowDatum> = {}
    const kidsByPath: Record<string, RowDatum[]> = {}
    for (const r of rows) {
      byPath[r.path.join('|')] = r
      const pk = r.path.slice(0, -1).join('|')
      ;(kidsByPath[pk] ||= []).push(r)
    }
    const leafId = field.slice(2)
    // Build path → rowNode index. No `getRowId` is defined, so
    // `api.getRowNode(...)` wouldn't work — iterate instead.
    const nodeByPath: Record<string, any> = {}
    api.forEachNode(n => {
      if (n.data) nodeByPath[(n.data.path as string[]).join('|')] = n
    })
    // Which Σ-columns include this leaf? We'll refresh & flash them on the
    // leaf row AND every ancestor row (the leaf itself also flashes for Σ
    // cols because the user sees a fresh aggregate value).
    const affectedSumCols: string[] = []
    for (const [colId, ids] of Object.entries(sumColLeavesRef.current)) {
      if (ids.includes(leafId)) affectedSumCols.push(colId)
    }
    const leafKey = leafPath.join('|')
    const leafNode = nodeByPath[leafKey]
    if (affectedSumCols.length > 0 && leafNode) {
      api.refreshCells({ rowNodes: [leafNode], columns: affectedSumCols, force: true })
      api.flashCells({
        rowNodes: [leafNode],
        columns: affectedSumCols,
        flashDuration: 1500,
        fadeDuration: 600,
      })
    }
    // Walk upward from leaf's parent to root.
    for (let depth = leafPath.length - 1; depth >= 1; depth--) {
      const parentPath = leafPath.slice(0, depth)
      const parent = byPath[parentPath.join('|')]
      if (!parent || parent.isLeaf) continue
      // Always propagate SUM upward for group rows as immediate feedback.
      // Server recalc will correct with proper formula values (weighted avg, etc.).
      const rule = parent[`__rule_${leafId}`]
      if (rule === 'empty') continue
      const kids = kidsByPath[parentPath.join('|')] || []
      let s = 0, has = false
      for (const k of kids) {
        const v = k[field]
        if (v == null || v === '') continue
        const n = parseFloat(String(v))
        if (!Number.isNaN(n)) { s += n; has = true }
      }
      const newVal = has ? String(s) : ''
      const changed = String(parent[field] ?? '') !== newVal
      parent[field] = newVal
      const node = nodeByPath[parentPath.join('|')]
      if (node) {
        // Refresh both the changed leaf-field column AND any Σ-columns that
        // include this leaf (they sum across periods on this row).
        const cols = [field, ...affectedSumCols]
        api.refreshCells({ rowNodes: [node], columns: cols, force: true })
        if (changed) {
          api.flashCells({
            rowNodes: [node],
            columns: cols,
            flashDuration: 1500,
            fadeDuration: 600,
          })
        }
      }
    }
  }, [])

  // ── Cell edit → save (non-blocking) ────────────────────────────────────
  const onCellValueChanged = useCallback((e: CellValueChangedEvent) => {
    const field: string | undefined = e.colDef.field
    if (!field || !field.startsWith('p_')) return
    const leafId = field.slice(2)
    const coordKey: string | undefined = e.data?.[`__coord_${leafId}`]
    if (!coordKey) return
    const newVal = e.newValue == null ? '' : String(e.newValue)
    // If the user overwrote a formula cell, flip its local rule to 'manual'
    // right away so the cell restyles (blue → yellow) without waiting for
    // the server round-trip.
    if (e.data[`__rule_${leafId}`] === 'formula') {
      e.data[`__rule_${leafId}`] = 'manual'
      e.node?.setData({ ...e.data })
    }
    // Optimistic: recompute parent sums locally for instant feedback.
    const path: string[] = e.data.path
    recomputeParentsForField(path, field)
    // Fire-and-forget save; server handles formula recalc. Don't block UI.
    setRecalcRunning(true)
    api.saveCells(sheetId, [{
      coord_key: coordKey,
      value: newVal,
      data_type: 'number',
      rule: 'manual',
      user_id: currentUserId,
    }]).then(() => {
      cellMapRef.current[coordKey] = newVal
      // Server recomputed formulas synchronously — refetch + flash so any
      // derived (formula/sum_children) cells that changed light up green.
      refreshAndFlashParents()
    }).catch(err => {
      setError(`Не удалось сохранить: ${err?.message || err}`)
    }).finally(() => {
      setRecalcRunning(false)
    })
  }, [sheetId, currentUserId, recomputeParentsForField, refreshAndFlashParents])

  // ── Context menu: chart + history ───────────────────────────────────
  const buildChartFromSelection = useCallback((chartType: string) => {
    const grid = gridApiRef.current
    if (!grid) return
    const ranges = grid.getCellRanges() || []
    if (ranges.length === 0) return
    const r = ranges[0]
    const cols = r.columns || []
    const si = Math.min(r.startRow?.rowIndex ?? 0, r.endRow?.rowIndex ?? 0)
    const ei = Math.max(r.startRow?.rowIndex ?? 0, r.endRow?.rowIndex ?? 0)
    const colLabels = cols.map(c => c.getColDef()?.headerName || c.getColId())
    const datasets: { label: string; data: number[] }[] = []
    for (let ri = si; ri <= ei; ri++) {
      const node = grid.getDisplayedRowAtIndex(ri)
      if (!node?.data) continue
      const rowLabel = node.data.label || `Row ${ri}`
      const values = cols.map(c => {
        const v = node.data[c.getColId()]
        return v != null ? (typeof v === 'number' ? v : parseFloat(String(v)) || 0) : 0
      })
      datasets.push({ label: rowLabel, data: values })
    }
    setChartOverlay({ type: chartType, labels: colLabels, datasets })
  }, [])

  // Render amCharts when overlay data changes
  useEffect(() => {
    if (!chartOverlay || !chartDivRef.current) return
    if (amRootRef.current) { amRootRef.current.dispose(); amRootRef.current = null }

    const root = am5.Root.new(chartDivRef.current)
    amRootRef.current = root
    root.setThemes([am5themes_Animated.new(root)])

    const { type, labels, datasets } = chartOverlay

    if (type === 'pie') {
      // Pie: flatten first dataset into category/value pairs
      const chart = root.container.children.push(
        am5percent.PieChart.new(root, { layout: root.verticalLayout })
      )
      const series = chart.series.push(
        am5percent.PieSeries.new(root, { valueField: 'value', categoryField: 'category', legendLabelText: '{category}', legendValueText: '{value}' })
      )
      series.labels.template.setAll({ fontSize: 12, text: '{category}: {valuePercentTotal.formatNumber("0.0")}%' })
      // For pie: if multiple rows selected, each row is a segment; if one row, each column is a segment
      let pieData: { category: string; value: number }[]
      if (datasets.length > 1) {
        // Multiple rows → each row becomes a pie segment (sum of its values)
        pieData = datasets.map(ds => ({ category: ds.label, value: ds.data.reduce((a, b) => a + b, 0) }))
      } else {
        // Single row → each column becomes a pie segment
        const d = datasets[0]?.data || []
        pieData = labels.map((l, i) => ({ category: l, value: d[i] || 0 }))
      }
      // Filter out zero/negative values for cleaner pie
      pieData = pieData.filter(d => d.value > 0)
      series.data.setAll(pieData)
      if (pieData.length > 1) {
        const legend = chart.children.push(am5.Legend.new(root, { centerX: am5.percent(50), x: am5.percent(50) }))
        legend.data.setAll(series.dataItems)
      }
      series.appear(1000, 100)
      chart.appear(1000, 100)
    } else {
      // XY chart (bar / line)
      const chart = root.container.children.push(
        am5xy.XYChart.new(root, { panX: true, panY: false, wheelX: 'panX', wheelY: 'zoomX', layout: root.verticalLayout })
      )
      const xRenderer = am5xy.AxisRendererX.new(root, { minGridDistance: 80 })
      xRenderer.labels.template.setAll({ rotation: -45, centerY: am5.percent(50), centerX: am5.percent(100), paddingRight: 8, fontSize: 11, oversizedBehavior: 'truncate', maxWidth: 120 })
      const xAxis = chart.xAxes.push(
        am5xy.CategoryAxis.new(root, { categoryField: 'category', renderer: xRenderer, tooltip: am5.Tooltip.new(root, {}) })
      )
      // Build merged data: one row per category with fields per series
      const data = labels.map((l, ci) => {
        const row: Record<string, any> = { category: l }
        for (const ds of datasets) row[ds.label] = ds.data[ci] || 0
        return row
      })
      xAxis.data.setAll(data)
      chart.yAxes.push(
        am5xy.ValueAxis.new(root, { renderer: am5xy.AxisRendererY.new(root, {}) })
      )
      for (const ds of datasets) {
        const xA = chart.xAxes.getIndex(0)!
        const yA = chart.yAxes.getIndex(0)!
        const s = type === 'bar'
          ? chart.series.push(am5xy.ColumnSeries.new(root, { name: ds.label, xAxis: xA, yAxis: yA, valueYField: ds.label, categoryXField: 'category', tooltip: am5.Tooltip.new(root, { labelText: '{name}: {valueY}' }) }))
          : chart.series.push(am5xy.LineSeries.new(root, { name: ds.label, xAxis: xA, yAxis: yA, valueYField: ds.label, categoryXField: 'category', tooltip: am5.Tooltip.new(root, { labelText: '{name}: {valueY}' }) }))
        s.data.setAll(data)
        s.appear(1000)
      }
      if (datasets.length > 1) {
        const legend = chart.children.push(am5.Legend.new(root, { centerX: am5.percent(50), x: am5.percent(50) }))
        legend.data.setAll(chart.series.values)
      }
      chart.set('cursor', am5xy.XYCursor.new(root, {}))
      chart.appear(1000, 100)
    }

    return () => { root.dispose(); amRootRef.current = null }
  }, [chartOverlay])

  const showHistory = useCallback(async (coordKey: string) => {
    setHistoryKey(coordKey)
    const data = await api.getCellHistory(sheetId, coordKey)
    setHistoryData(data)
    setHistoryOpen(true)
  }, [sheetId])

  const getContextMenuItems = useCallback((params: any): any[] => {
    const node = params.node
    const colId = params.column?.getColId?.() || ''
    const coordKey = node?.data?.[`__coord_${colId.replace('p_', '')}`] || ''
    return [
      'copy', 'copyWithHeaders', 'paste', 'separator', 'export',
      'separator',
      {
        name: 'История изменений',
        disabled: !coordKey,
        action: () => { if (coordKey) showHistory(coordKey) },
      },
      'separator',
      {
        name: 'График',
        subMenu: [
          { name: 'Столбчатая', action: () => buildChartFromSelection('bar') },
          { name: 'Линейная', action: () => buildChartFromSelection('line') },
          { name: 'Круговая', action: () => buildChartFromSelection('pie') },
        ],
      },
    ]
  }, [showHistory, buildChartFromSelection])

  // ── Render ─────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: 2, minHeight: 200 }}>
        <CircularProgress size={48} thickness={3} sx={{ color: 'primary.main' }} />
        <Typography variant="body2" color="text.secondary">Загрузка…</Typography>
      </Box>
    )
  }
  if (error) {
    return <Box sx={{ p: 2, color: 'error.main' }}>Ошибка: {error}</Box>
  }

  const pinnedEntries = Object.keys(pinned).filter(aId => !!pinned[aId])

  return (
    <Box sx={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column' }}>
      {/* Toolbar / drop zone for analytic pinning. Always visible so users can
          drop rows here; shows hint when empty. */}
      <Box
        sx={{
          display: 'flex', alignItems: 'center', gap: 0.5, flexWrap: 'wrap',
          px: 1, py: 0.5, borderBottom: '1px solid #f0f0f0', minHeight: 32,
          bgcolor: '#fafafa',
        }}
        onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy' }}
        onDrop={e => {
          e.preventDefault()
          try {
            const data = JSON.parse(e.dataTransfer.getData('text/plain'))
            if (data.analyticId && data.recordId) handlePin(data.analyticId, data.recordId)
          } catch { /* ignore */ }
        }}
      >
        {pinnedEntries.length === 0 && (
          <Typography sx={{ fontSize: 11, color: '#999' }}>
            Перетащите строку сюда, чтобы зафиксировать аналитику
          </Typography>
        )}
        {pinnedEntries.map(aId => {
          const aName = analyticsMap[aId]?.name || aId.slice(0, 6)
          const rName = recordNames[pinned[aId]] || '?'
          return (
            <Chip
              key={aId}
              size="small"
              label={`${aName}: ${rName}`}
              onDelete={() => handleUnpin(aId)}
              sx={{ fontSize: 12 }}
            />
          )
        })}
        {/* Column-level aggregate toggles (Годы / Кварталы).
            3-state cycle: end → start → hidden → end.
            'end' = show after leaves (default), 'start' = show before, 'hidden' = off. */}
        {colLevelNames.length > 0 && (
          <Box sx={{ display: 'flex', gap: 0.5, ml: pinnedEntries.length ? 1 : 0 }}>
            {colLevelNames.map(({ level, label }) => {
              const state = colLevelToggles[level] || 'hidden'
              const nextState = (s: ColLevelState): ColLevelState =>
                s === 'end' ? 'start' : s === 'start' ? 'hidden' : 'end'
              const icon = state === 'start' ? '◀' : state === 'end' ? '▶' : ''
              return (
                <Chip
                  key={`lvl-${level}`}
                  size="small"
                  label={`Σ ${label}${icon ? ' ' + icon : ''}`}
                  color={state !== 'hidden' ? 'primary' : 'default'}
                  variant={state !== 'hidden' ? 'filled' : 'outlined'}
                  onClick={() => setColLevelToggles(prev => ({ ...prev, [level]: nextState(state) }))}
                  sx={{ fontSize: 11 }}
                  data-testid={`col-level-chip-${level}`}
                />
              )
            })}
          </Box>
        )}
      </Box>
      <Box sx={{ flex: 1, minHeight: 0 }} ref={gridContainerRef}>
        <AgGridReact
          theme={themeAlpineWithBorders}
          rowData={rowData}
          columnDefs={columnDefs}
          treeData
          getDataPath={getDataPath}
          autoGroupColumnDef={autoGroupColumnDef}
          groupDefaultExpanded={-1}
          rowHeight={28}
          headerHeight={30}
          onCellValueChanged={onCellValueChanged}
          onCellKeyDown={onCellKeyDown}
          onGridReady={onGridReady}
          onColumnResized={e => { if (e.finished) captureColumnState() }}
          onColumnMoved={e => { if (e.finished) captureColumnState() }}
          onColumnPinned={captureColumnState}
          onColumnVisible={captureColumnState}
          tooltipShowDelay={1000}
          animateRows={false}
          stopEditingWhenCellsLoseFocus
          singleClickEdit={false}
          enterNavigatesVertically
          enterNavigatesVerticallyAfterEdit
          cellSelection={{ handle: { mode: 'fill' } }}
          undoRedoCellEditing
          undoRedoCellEditingLimit={50}
          suppressRowHoverHighlight
          rowSelection={undefined}
          onCellSelectionChanged={updateSelectionStats}
          getContextMenuItems={getContextMenuItems}
          statusBar={{
            statusPanels: [
              { statusPanel: 'recalcStatusPanel', align: 'left' },
              { statusPanel: 'selectionStatusPanel', align: 'right' },
            ],
          }}
          components={{
            recalcStatusPanel: RecalcStatusPanel,
            selectionStatusPanel: SelectionStatusPanel,
          }}
          defaultColDef={{
            resizable: true,
            sortable: false,
            filter: false,
            headerClass: 'ag-center-header',
            // Commit edit on arrow keys — AG Grid will navigate to the next
            // cell naturally after stopEditing.
            suppressKeyboardEvent: (params: any) => {
              if (!params.editing) return false
              const k = params.event?.key
              if (k === 'ArrowUp' || k === 'ArrowDown' || k === 'ArrowLeft' || k === 'ArrowRight') {
                params.api.stopEditing(false)
                return false // let grid consume the arrow for navigation
              }
              return false
            },
          }}
        />
      </Box>

      <FormulaEditor
        open={formulaEditorOpen}
        formula={formulaMapRef.current[formulaEditorKey] || ''}
        rule={cellRuleRef.current[formulaEditorKey] || 'manual'}
        modelId={modelId}
        currentSheetId={sheetId}
        onClose={() => setFormulaEditorOpen(false)}
        onSave={async (text, rule) => {
          try {
            await api.saveCells(sheetId, [{
              coord_key: formulaEditorKey,
              formula: text,
              rule: rule as any,
              user_id: currentUserId,
            }])
            formulaMapRef.current[formulaEditorKey] = text
            cellRuleRef.current[formulaEditorKey] = rule as any
            // Offer to promote this per-cell formula to a rule on the indicator
            // (P3 snackbar). Requires we know the main analytic — extract
            // indicator_id from coord_key by slicing at its index in dbOrd.
            const idx = mainAnalyticIdxRef.current
            const parts = formulaEditorKey.split('|')
            const indicatorId = idx >= 0 && idx < parts.length ? parts[idx] : ''
            if (indicatorId) {
              setPromoteSnack({ coordKey: formulaEditorKey, formula: text, indicatorId })
            }
            // Fetch only the updated cell(s) instead of reloading the entire grid.
            const freshCells = await api.getCellsPartial(sheetId, [formulaEditorKey], currentUserId)
            for (const c of freshCells) {
              cellMapRef.current[c.coord_key] = c.value ?? ''
              if (c.rule) cellRuleRef.current[c.coord_key] = c.rule as CellRule
              if (c.formula) formulaMapRef.current[c.coord_key] = c.formula
            }
            // Find the period column from coord_key and update the row node in-place.
            const colAIdx = dbOrdRef.current.indexOf(colAIdRef.current)
            const periodRecId = colAIdx >= 0 && colAIdx < parts.length ? parts[colAIdx] : ''
            const colId = periodRecId ? `p_${periodRecId}` : ''
            const grid = gridApiRef.current
            if (grid && colId) {
              const newVal = cellMapRef.current[formulaEditorKey] ?? ''
              grid.forEachNode(node => {
                if (node.data?.[`__coord_${periodRecId}`] === formulaEditorKey) {
                  node.data[colId] = newVal
                  grid.refreshCells({ rowNodes: [node], columns: [colId], force: true })
                  grid.flashCells({ rowNodes: [node], columns: [colId] })
                }
              })
            }
          } catch (err: any) {
            setError(`Не удалось сохранить формулу: ${err?.message || err}`)
          }
        }}
      />

      {/* Chart overlay */}
      {chartOverlay && (
        <Box sx={{
          position: 'absolute', inset: 40, zIndex: 1300,
          bgcolor: '#fff', border: '1px solid #e0e0e0', borderRadius: 2,
          boxShadow: 4, display: 'flex', flexDirection: 'column',
        }}>
          <Box sx={{ display: 'flex', justifyContent: 'flex-end', p: 0.5 }}>
            <IconButton size="small" onClick={() => {
              if (amRootRef.current) { amRootRef.current.dispose(); amRootRef.current = null }
              setChartOverlay(null)
            }}>
              <CloseOutlined fontSize="small" />
            </IconButton>
          </Box>
          <div ref={chartDivRef} style={{ flex: 1, minHeight: 0, width: '100%' }} />
        </Box>
      )}

      {/* History dialog */}
      <Dialog open={historyOpen} onClose={() => setHistoryOpen(false)} maxWidth="sm" fullWidth>
        <DialogTitle sx={{ py: 1 }}>История изменений</DialogTitle>
        <DialogContent>
          {historyData.length === 0 ? (
            <Typography variant="body2" color="textSecondary">Нет изменений</Typography>
          ) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
              <thead>
                <tr>
                  <th style={{ textAlign: 'left', padding: '4px 8px', borderBottom: '1px solid #e0e0e0' }}>Дата/время</th>
                  <th style={{ textAlign: 'left', padding: '4px 8px', borderBottom: '1px solid #e0e0e0' }}>Пользователь</th>
                  <th style={{ textAlign: 'right', padding: '4px 8px', borderBottom: '1px solid #e0e0e0' }}>Было</th>
                  <th style={{ textAlign: 'right', padding: '4px 8px', borderBottom: '1px solid #e0e0e0' }}>Стало</th>
                </tr>
              </thead>
              <tbody>
                {historyData.map((h: any, i: number) => (
                  <tr key={i}>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid #f0f0f0' }}>{h.created_at?.replace('T', ' ').slice(0, 19)}</td>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid #f0f0f0' }}>{h.username || '—'}</td>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid #f0f0f0', textAlign: 'right', color: '#999' }}>{h.old_value ?? '—'}</td>
                    <td style={{ padding: '4px 8px', borderBottom: '1px solid #f0f0f0', textAlign: 'right' }}>{h.new_value ?? '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </DialogContent>
      </Dialog>

      <Snackbar
        open={!!promoteSnack}
        autoHideDuration={8000}
        onClose={() => setPromoteSnack(null)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
        message="Формула сохранена на клетку"
        action={
          <>
            <MUIButton
              color="primary"
              size="small"
              onClick={async () => {
                if (!promoteSnack) return
                try {
                  await api.promoteCellToRule(
                    sheetId,
                    promoteSnack.indicatorId,
                    promoteSnack.coordKey,
                    promoteSnack.formula,
                  )
                  setPromoteSnack(null)
                  load()
                } catch (err: any) {
                  setError(`Не удалось сделать правилом: ${err?.message || err}`)
                }
              }}
            >
              Сделать правилом показателя
            </MUIButton>
            <MUIButton color="inherit" size="small" onClick={() => setPromoteSnack(null)}>
              Закрыть
            </MUIButton>
          </>
        }
      />
    </Box>
  )
}
