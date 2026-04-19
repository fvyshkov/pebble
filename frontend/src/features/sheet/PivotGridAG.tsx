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
import { Box, Chip, CircularProgress, LinearProgress, Tooltip, Typography } from '@mui/material'
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
import {
  AllEnterpriseModule,
  LicenseManager,
} from 'ag-grid-enterprise'
import * as api from '../../api'
import type { SheetAnalytic, Analytic, AnalyticRecord, CellData } from '../../types'

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
export default function PivotGridAG({ sheetId, currentUserId, calcProgress }: Props) {
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
  const [analyticsMap, setAnalyticsMap] = useState<Record<string, Analytic>>({})
  const [recordNames, setRecordNames] = useState<Record<string, string>>({})
  const vsLoadedRef = useRef(false)
  // Persisted AG Grid column state (widths, order, pinned, hidden) — declared
  // here so the debounced save effect can depend on `columnStateVersion`.
  const savedColumnStateRef = useRef<any[] | null>(null)
  const [columnStateVersion, setColumnStateVersion] = useState(0)

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
  const colLeafIdsRef = useRef<string[]>([])
  const dbOrdRef = useRef<string[]>([])
  const colAIdRef = useRef<string>('')
  const pinnedRef = useRef<Record<string, string>>({})
  const analyticsRef = useRef<Record<string, Analytic>>({})

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
      await Promise.all(sa.map(async b => {
        const [analytic, recs] = await Promise.all([
          api.getAnalytic(b.analytic_id),
          api.listRecords(b.analytic_id),
        ])
        aMap[b.analytic_id] = analytic
        rMap[b.analytic_id] = buildRecordTree(recs)
        for (const r of recs) {
          const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
          recNameMap[r.id] = (d && d.name) || r.id.slice(0, 6)
        }
      }))
      analyticsRef.current = aMap
      setAnalyticsMap(aMap)
      setRecordNames(recNameMap)

      const dbOrd = sa.map(b => b.analytic_id)
      dbOrdRef.current = dbOrd

      // View settings — read once per mount. Subsequent reloads reuse state.
      let curOrder = [...dbOrd]
      let effectivePinned = pinned
      if (!vsLoadedRef.current) {
        try {
          const vs: any = await api.getViewSettings(sheetId)
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

      // ── Column definitions ──────────────────────────────────────────────
      const buildColDefs = (nodes: RecordNode[]): (ColDef | ColGroupDef)[] =>
        nodes.map(n => {
          const label = recordLabel(n)
          if (n.children.length === 0) {
            return makePeriodColDef(label, n.record.id)
          }
          return {
            headerName: label,
            headerClass: 'ag-center-header',
            children: buildColDefs(n.children),
          } as ColGroupDef
        })

      // ── Cells (full prefetch) ───────────────────────────────────────────
      cellMapRef.current = {}
      cellRuleRef.current = {}
      const cellData: CellData[] = await api.getCells(sheetId, currentUserId)
      for (const c of cellData) {
        cellMapRef.current[c.coord_key] = c.value ?? ''
        if (c.rule) cellRuleRef.current[c.coord_key] = c.rule as CellRule
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
              }
              if (parts.length >= 2) {
                const lookup = lookupCell(parts)
                row[`p_${leaf.record.id}`] = lookup.value
                row[`__coord_${leaf.record.id}`] = parts.join('|')
                row[`__parts_${leaf.record.id}`] = parts
                // Legacy resolveRule: group rows default to sum_children,
                // terminal rows to manual, unless backend stored explicit rule.
                const storedRule = cellRuleRef.current[parts.join('|')]
                row[`__rule_${leaf.record.id}`] = storedRule
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
      const defs: (ColDef | ColGroupDef)[] = buildColDefs(colTree)
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
      }).catch(() => { /* ignore */ })
    }, 500)
    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current)
    }
  }, [pinned, columnStateVersion, sheetId])

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

  // When a model-wide recalc ends, refetch cells and apply diffs in place so
  // AG Grid's enableCellChangeFlash lights up every cell that actually
  // changed (green, per the CSS override). Works regardless of WHICH user
  // triggered the recalc — we always diff the fresh server state against
  // our local snapshot.
  const prevCalcProgressRef = useRef(calcProgress)
  useEffect(() => {
    const wasRunning = !!prevCalcProgressRef.current
    const nowIdle = !calcProgress
    prevCalcProgressRef.current = calcProgress
    if (!wasRunning || !nowIdle) return
    const grid = gridApiRef.current
    if (!grid) return
    let cancelled = false
    ;(async () => {
      try {
        const fresh: CellData[] = await api.getCells(sheetId, currentUserId)
        if (cancelled) return
        // Build fresh maps.
        const freshMap: Record<string, string> = {}
        const freshRule: Record<string, CellRule> = {}
        for (const c of fresh) {
          freshMap[c.coord_key] = c.value ?? ''
          if (c.rule) freshRule[c.coord_key] = c.rule as CellRule
        }
        // Walk every displayed leaf row; record every (row, col) whose
        // DERIVED value (formula / sum_children, NOT manual) changed so
        // we can flash it. Manual cells aren't flashed — they're the input
        // that triggered the recalc, not a result of it.
        const flashRowNodes: any[] = []
        const flashColSet = new Set<string>()
        grid.forEachNode(node => {
          if (!node.data || !node.data.isLeaf) return
          let rowHasDerivedChange = false
          for (const leafId of colLeafIdsRef.current) {
            const coordKey: string | undefined = node.data[`__coord_${leafId}`]
            if (!coordKey) continue
            const newVal = freshMap[coordKey] ?? ''
            const oldVal = (node.data[`p_${leafId}`] ?? '') as string | number
            // Keep rule map fresh too (formula→manual transitions etc.).
            const newRule = (freshRule[coordKey] as CellRule) || 'manual'
            const oldRule: CellRule = node.data[`__rule_${leafId}`] || 'manual'
            if (newRule !== oldRule) node.data[`__rule_${leafId}`] = newRule
            if (String(newVal) !== String(oldVal)) {
              node.setDataValue(`p_${leafId}`, newVal)
              if (newRule !== 'manual') {
                rowHasDerivedChange = true
                flashColSet.add(`p_${leafId}`)
              }
            }
          }
          if (rowHasDerivedChange) flashRowNodes.push(node)
        })
        if (flashRowNodes.length > 0 && flashColSet.size > 0) {
          grid.flashCells({
            rowNodes: flashRowNodes,
            columns: Array.from(flashColSet),
            flashDuration: 1500,
            fadeDuration: 600,
          })
        }
        // Update refs for future lookups/diffs.
        cellMapRef.current = freshMap
        cellRuleRef.current = freshRule
      } catch { /* ignore — next edit/recalc will retry */ }
    })()
    return () => { cancelled = true }
  }, [calcProgress, sheetId, currentUserId])

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
    // Rough initial estimate (AG Grid's autoSizeColumns on first-data-rendered
    // refines this). Keep it generous so headers aren't clipped pre-measurement.
    const chars = (label || '').length
    const autoWidth = Math.max(90, Math.min(260, chars * 10 + 28))
    return {
      headerName: label,
      field,
      headerClass: 'ag-center-header',
      width: autoWidth,
      minWidth: 70,
      // NB: do NOT set enableCellChangeFlash here — we only want to flash
      // DERIVED cells (sums/formulas that recomputed), not the manual cell
      // the user just typed into. Flash is triggered explicitly via
      // api.flashCells() in the recalc-diff effect and in the local
      // parent-sum update path.
      editable: (p: any) => {
        if (!p.data || !p.data.isLeaf) return false // group rows: read-only
        const rule: CellRule = p.data[`__rule_${periodRecId}`] || 'manual'
        return rule === 'manual'
      },
      valueFormatter: (p: any) => {
        const rule: CellRule = p.data?.[`__rule_${periodRecId}`] || 'manual'
        if (rule === 'empty') return ''
        const v = p.value
        if (v == null || v === '') return ''
        const num = Number(v)
        if (!Number.isNaN(num)) {
          return num.toLocaleString('ru-RU', { maximumFractionDigits: 2 })
        }
        return String(v)
      },
      valueParser: (p: any) => {
        // Reject non-numeric input for numeric cells.
        const s = String(p.newValue ?? '').trim().replace(',', '.')
        if (s === '') return ''
        const num = Number(s)
        if (Number.isNaN(num)) return p.oldValue
        return s
      },
      cellStyle: (p: any) => {
        const isLeaf = !!p.data?.isLeaf
        const rule: CellRule = p.data?.[`__rule_${periodRecId}`] || 'manual'
        const s: any = { textAlign: 'right' }
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
    }
  }

  const getDataPath = useCallback((data: RowDatum) => data.path, [])

  const autoGroupColumnDef = useMemo<ColDef>(() => ({
    headerName: 'Аналитика',
    minWidth: 280,
    pinned: 'left',
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
      } else {
        const colIds = g.getAllDisplayedColumns()
          .map(c => c.getColId())
          .filter(id => id !== 'ag-Grid-AutoColumn')
        if (colIds.length > 0) g.autoSizeColumns(colIds, false)
      }
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
          if (rule === 'manual') {
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
    // Walk upward from leaf's parent to root.
    for (let depth = leafPath.length - 1; depth >= 1; depth--) {
      const parentPath = leafPath.slice(0, depth)
      const parent = byPath[parentPath.join('|')]
      if (!parent) continue
      const rule = parent[`__rule_${leafId}`]
      if (rule !== 'sum_children') continue
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
      // Nudge AG Grid to refresh just that cell.
      const node = api.getRowNode(parentPath.join('|'))
      if (node) {
        api.refreshCells({ rowNodes: [node], columns: [field], force: true })
        // Flash the recomputed parent sum cell (it's a derived value).
        if (changed) {
          api.flashCells({
            rowNodes: [node],
            columns: [field],
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
    }).catch(err => {
      setError(`Не удалось сохранить: ${err?.message || err}`)
    }).finally(() => {
      setRecalcRunning(false)
    })
  }, [sheetId, currentUserId, recomputeParentsForField])

  // ── Render ─────────────────────────────────────────────────────────────
  if (loading) {
    return (
      <Box sx={{ p: 3, display: 'flex', gap: 1, alignItems: 'center' }}>
        <CircularProgress size={18} />
        <Typography variant="body2">Загрузка AG Grid…</Typography>
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
      </Box>
      <Box sx={{ flex: 1, minHeight: 0 }} ref={gridContainerRef}>
        <AgGridReact
          theme={themeAlpine}
          rowData={rowData}
          columnDefs={columnDefs}
          treeData
          getDataPath={getDataPath}
          autoGroupColumnDef={autoGroupColumnDef}
          groupDefaultExpanded={0}
          rowHeight={28}
          headerHeight={30}
          onCellValueChanged={onCellValueChanged}
          onCellKeyDown={onCellKeyDown}
          onGridReady={onGridReady}
          onColumnResized={e => { if (e.finished) captureColumnState() }}
          onColumnMoved={e => { if (e.finished) captureColumnState() }}
          onColumnPinned={captureColumnState}
          onColumnVisible={captureColumnState}
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
    </Box>
  )
}
