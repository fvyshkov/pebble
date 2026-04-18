/**
 * AG Grid-based pivot view (Phase 1).
 *
 * Shows the same data as the custom PivotGrid but using AG Grid Community +
 * Enterprise (row grouping, column groups, fill handle, keyboard nav).
 * Parity todo list is tracked in tests/plan_aggrid.md.
 */
import { useEffect, useMemo, useState, useCallback, useRef } from 'react'
import { Box, CircularProgress, Typography } from '@mui/material'
import { AgGridReact } from 'ag-grid-react'
import {
  ModuleRegistry,
  AllCommunityModule,
  themeAlpine,
  type ColDef,
  type ColGroupDef,
  type CellValueChangedEvent,
} from 'ag-grid-community'
import {
  AllEnterpriseModule,
  LicenseManager,
} from 'ag-grid-enterprise'
import * as api from '../../api'
import type { SheetAnalytic, Analytic, AnalyticRecord } from '../../types'

ModuleRegistry.registerModules([AllCommunityModule, AllEnterpriseModule])
// Unlicensed Enterprise = watermark on-screen. Acceptable for dev/preview.
// If you buy a license, put the key in VITE_AG_GRID_LICENSE.
const licenseKey = (import.meta as any).env?.VITE_AG_GRID_LICENSE
if (licenseKey) LicenseManager.setLicenseKey(licenseKey)

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

function recordLabel(n: RecordNode, _analytic: Analytic | undefined): string {
  return (n.data && n.data.name) || n.record.id.slice(0, 8)
}

// ── Props ──────────────────────────────────────────────────────────────────
interface Props {
  sheetId: string
  modelId: string
  currentUserId?: string
}

interface RowAnalyticKey {
  [aId: string]: string // recordId
}

// ── Component ──────────────────────────────────────────────────────────────
export default function PivotGridAG({ sheetId, currentUserId }: Props) {
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [rowData, setRowData] = useState<any[]>([])
  const [columnDefs, setColumnDefs] = useState<(ColDef | ColGroupDef)[]>([])
  const [sheetName, setSheetName] = useState('')
  // Keep fresh copies to avoid stale closure in editor
  const analyticsRef = useRef<Record<string, Analytic>>({})
  const rowAIdsRef = useRef<string[]>([])
  const dbOrdRef = useRef<string[]>([])

  // ── Data load ────────────────────────────────────────────────────────────
  const load = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      const sa: SheetAnalytic[] = await api.listSheetAnalytics(sheetId)
      if (sa.length < 2) {
        throw new Error('Нужно минимум 2 аналитики на листе (колонки + строки)')
      }
      const aMap: Record<string, Analytic> = {}
      const rMap: Record<string, RecordNode[]> = {}
      await Promise.all(sa.map(async b => {
        const [analytic, recs] = await Promise.all([
          api.getAnalytic(b.analytic_id),
          api.listRecords(b.analytic_id),
        ])
        aMap[b.analytic_id] = analytic
        rMap[b.analytic_id] = buildRecordTree(recs)
      }))
      analyticsRef.current = aMap

      const dbOrd = sa.map(b => b.analytic_id)
      dbOrdRef.current = dbOrd

      // View settings for order + pinned
      let curOrder = [...dbOrd]
      let pinned: Record<string, string> = {}
      try {
        const vs: any = await api.getViewSettings(sheetId)
        if (vs?.order?.length) {
          const valid = new Set(dbOrd)
          curOrder = vs.order.filter((id: string) => valid.has(id))
          for (const id of dbOrd) if (!curOrder.includes(id)) curOrder.push(id)
        }
        if (vs?.pinned) pinned = vs.pinned
      } catch { /* ignore */ }

      const colAId = curOrder[0]
      const rowAIds = curOrder.slice(1).filter(id => !pinned[id])
      rowAIdsRef.current = rowAIds

      // Column leaves (periods) from first analytic
      const colTree = rMap[colAId] || []
      const colLeaves = getLeaves(colTree)

      // Build column definitions — mirror period hierarchy as column groups
      const colAnalytic = aMap[colAId]
      const buildColDefs = (nodes: RecordNode[]): (ColDef | ColGroupDef)[] =>
        nodes.map(n => {
          const label = recordLabel(n, colAnalytic)
          if (n.children.length === 0) {
            return {
              headerName: label,
              field: `p_${n.record.id}`,
              editable: true,
              minWidth: 110,
              valueFormatter: (p: any) => {
                const v = p.value
                if (v == null || v === '') return ''
                const num = Number(v)
                if (!Number.isNaN(num)) {
                  return num.toLocaleString('ru-RU', { maximumFractionDigits: 2 })
                }
                return String(v)
              },
              cellStyle: (p: any) => {
                const s: any = { textAlign: 'right' }
                const v = p.value
                const num = Number(v)
                if (!Number.isNaN(num) && num < 0) s.color = '#d32f2f'
                return s
              },
            } as ColDef
          }
          return {
            headerName: label,
            children: buildColDefs(n.children),
          } as ColGroupDef
        })

      // Row-group columns: one per row analytic
      const rowGroupCols: ColDef[] = rowAIds.map((aId, idx) => ({
        headerName: aMap[aId]?.name || 'Строка',
        field: `r_${aId}`,
        rowGroup: idx < rowAIds.length - 1,
        hide: idx < rowAIds.length - 1,
        minWidth: 200,
        pinned: 'left',
      }))

      const defs: (ColDef | ColGroupDef)[] = [
        ...rowGroupCols,
        ...buildColDefs(colTree),
      ]
      setColumnDefs(defs)

      // ── Cells ──
      const cellData = await api.getCells(sheetId, currentUserId)
      const cellMap: Record<string, string> = {}
      for (const c of cellData) cellMap[c.coord_key] = c.value ?? ''

      // ── Row data: cartesian product of row leaves ──
      // Build combinations of leaf records for each row analytic (ignoring hierarchy —
      // AG Grid handles grouping from the r_<aId> fields).
      const rowLeaves: RecordNode[][] = rowAIds.map(aId => getLeaves(rMap[aId] || []))
      const cartesian = (lists: RecordNode[][]): RecordNode[][] => {
        if (lists.length === 0) return [[]]
        const [head, ...rest] = lists
        const restProd = cartesian(rest)
        const out: RecordNode[][] = []
        for (const h of head) for (const r of restProd) out.push([h, ...r])
        return out
      }

      const combos = cartesian(rowLeaves)
      const rows: any[] = []
      for (const combo of combos) {
        const rowRecordIds: RowAnalyticKey = {}
        rowAIds.forEach((aId, i) => { rowRecordIds[aId] = combo[i].record.id })

        const row: any = {}
        rowAIds.forEach((aId, i) => {
          row[`r_${aId}`] = recordLabel(combo[i], aMap[aId])
        })

        for (const leaf of colLeaves) {
          // coord_key = join of record ids in dbOrd order
          const parts: string[] = []
          for (const aId of dbOrd) {
            if (aId === colAId) parts.push(leaf.record.id)
            else if (pinned[aId]) parts.push(pinned[aId])
            else if (rowRecordIds[aId]) parts.push(rowRecordIds[aId])
          }
          const coordKey = parts.join('|')
          row[`p_${leaf.record.id}`] = cellMap[coordKey] ?? ''
          row[`__coord_${leaf.record.id}`] = coordKey
        }
        rows.push(row)
      }
      setRowData(rows)
      setSheetName(sa[0] ? `${rowAIds.length} аналитик × ${colLeaves.length} периодов` : '')
      setLoading(false)
    } catch (e: any) {
      setError(e?.message || String(e))
      setLoading(false)
    }
  }, [sheetId, currentUserId])

  useEffect(() => { load() }, [load])

  // ── Cell edit → save ─────────────────────────────────────────────────────
  const onCellValueChanged = useCallback(async (e: CellValueChangedEvent) => {
    const field: string | undefined = e.colDef.field
    if (!field || !field.startsWith('p_')) return
    const leafId = field.slice(2)
    const coordKey: string | undefined = e.data?.[`__coord_${leafId}`]
    if (!coordKey) return
    const newVal = e.newValue == null ? '' : String(e.newValue)
    try {
      await api.saveCells(sheetId, [{
        coord_key: coordKey,
        value: newVal,
        data_type: 'number',
        rule: 'manual',
        user_id: currentUserId,
      }])
    } catch (err: any) {
      setError(`Не удалось сохранить: ${err?.message || err}`)
    }
  }, [sheetId, currentUserId])

  // ── Render ───────────────────────────────────────────────────────────────
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

  return (
    <Box sx={{ width: '100%', height: '100%', display: 'flex', flexDirection: 'column' }}>
      <Box sx={{ px: 1.5, py: 0.5, fontSize: 12, color: '#666' }}>
        AG Grid · {sheetName}
      </Box>
      <Box sx={{ flex: 1, minHeight: 0 }}>
        <AgGridReact
          theme={themeAlpine}
          rowData={rowData}
          columnDefs={columnDefs}
          onCellValueChanged={onCellValueChanged}
          animateRows={false}
          stopEditingWhenCellsLoseFocus
          singleClickEdit={false}
          enterNavigatesVertically
          enterNavigatesVerticallyAfterEdit
          enableRangeSelection
          enableFillHandle
          undoRedoCellEditing
          undoRedoCellEditingLimit={50}
          suppressRowHoverHighlight={false}
          autoGroupColumnDef={{
            minWidth: 260,
            pinned: 'left',
            cellRendererParams: { suppressCount: true },
          }}
          groupDisplayType="multipleColumns"
          defaultColDef={{
            resizable: true,
            sortable: false,
            filter: false,
          }}
        />
      </Box>
    </Box>
  )
}
