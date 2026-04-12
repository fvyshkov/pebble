import { useState, useEffect, useCallback, useMemo, useRef } from 'react'
import {
  Box, Typography, IconButton, Tooltip, Dialog, DialogTitle, DialogContent,
  List, ListItem, ListItemIcon, ListItemText, Chip, Popover,
  Select, MenuItem,
} from '@mui/material'
import FormatListNumberedOutlined from '@mui/icons-material/FormatListNumberedOutlined'
import FileDownloadOutlined from '@mui/icons-material/FileDownloadOutlined'
import FileUploadOutlined from '@mui/icons-material/FileUploadOutlined'
import CalculateOutlined from '@mui/icons-material/CalculateOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import PushPinOutlined from '@mui/icons-material/PushPinOutlined'
import * as Icons from '@mui/icons-material'
import MoreHorizOutlined from '@mui/icons-material/MoreHorizOutlined'
import * as api from '../../api'
import type { SheetAnalytic, Analytic, AnalyticRecord } from '../../types'
import FormulaEditor from './FormulaEditor'
// Formula evaluation is now fully server-side

// ─── Tree helpers ───
interface RecordNode {
  record: AnalyticRecord; data: Record<string, any>; children: RecordNode[]
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
  const r: RecordNode[] = []
  const w = (ns: RecordNode[]) => { for (const n of ns) n.children.length === 0 ? r.push(n) : w(n.children) }
  w(nodes); return r
}
function leafCount(n: RecordNode): number { return n.children.length === 0 ? 1 : n.children.reduce((s, c) => s + leafCount(c), 0) }
function maxDepth(ns: RecordNode[]): number { return ns.length === 0 ? 0 : 1 + Math.max(0, ...ns.map(n => maxDepth(n.children))) }
function buildHeaderRows(tree: RecordNode[]) {
  const depth = maxDepth(tree); if (depth === 0) return []
  const rows: { node: RecordNode; colspan: number; rowspan: number }[][] = Array.from({ length: depth }, () => [])
  const walk = (nodes: RecordNode[], lvl: number) => {
    for (const n of nodes) {
      rows[lvl].push({ node: n, colspan: leafCount(n), rowspan: n.children.length === 0 ? depth - lvl : 1 })
      if (n.children.length > 0) walk(n.children, lvl + 1)
    }
  }
  walk(tree, 0); return rows
}
function flattenWithLevel(nodes: RecordNode[], lvl = 0): { node: RecordNode; level: number }[] {
  const r: { node: RecordNode; level: number }[] = []
  for (const n of nodes) { r.push({ node: n, level: lvl }); if (n.children.length > 0) r.push(...flattenWithLevel(n.children, lvl + 1)) }
  return r
}
function flattenWithLevelAndParents(nodes: RecordNode[], lvl = 0, parents: string[] = []): { node: RecordNode; level: number; parentChain: string[] }[] {
  const r: { node: RecordNode; level: number; parentChain: string[] }[] = []
  for (const n of nodes) {
    r.push({ node: n, level: lvl, parentChain: [...parents] })
    if (n.children.length > 0) r.push(...flattenWithLevelAndParents(n.children, lvl + 1, [...parents, n.record.id]))
  }
  return r
}
function findNodeById(nodes: RecordNode[], id: string): RecordNode | null {
  for (const n of nodes) { if (n.record.id === id) return n; const f = findNodeById(n.children, id); if (f) return f }
  return null
}

// ─── Formatting ───
function unitToDataType(unit: string | undefined, fallback: string): string {
  if (!unit) return fallback
  const u = unit.toLowerCase().trim()
  if (u === 'шт' || u === 'шт.' || u === 'мес' || u === 'мес.') return 'quantity'
  if (u === '%') return 'percent'
  return fallback
}

function fmtDisplay(val: string | undefined, dt: string): string {
  if (!val || val === '') return ''
  if (dt === 'string') return val
  const num = parseFloat(val)
  if (isNaN(num)) return val
  if (dt === 'sum') return num.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  if (dt === 'percent') return (num * 100).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%'
  if (dt === 'quantity') return Math.round(num).toLocaleString('ru-RU')
  return val
}

// ─── Cell with edit/display modes ───
function PivotCell({ value, onChange, dataType, editable, forceEdit, onStopEdit }: {
  value: string; onChange: (v: string) => void; dataType: string; editable: boolean
  forceEdit?: boolean; onStopEdit?: () => void
}) {
  const [editing, setEditing] = useState(false)
  const [local, setLocal] = useState(value)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => { setLocal(value) }, [value])
  useEffect(() => { if (editing) inputRef.current?.focus() }, [editing])

  // Parent requests edit via Enter key
  useEffect(() => { if (forceEdit && editable && !editing) setEditing(true) }, [forceEdit])

  if (!editable) {
    return (
      <div style={{
        padding: '4px 6px', fontSize: 13, textAlign: dataType === 'string' ? 'left' : 'right',
        background: '#fff', color: '#666', minHeight: 24, userSelect: 'none',
      }}>
        {fmtDisplay(value, dataType)}
      </div>
    )
  }

  if (!editing) {
    return (
      <div
        data-editable-cell
        onClick={() => setEditing(true)}
        onDoubleClick={() => setEditing(true)}
        style={{
          padding: '4px 6px', fontSize: 13, textAlign: dataType === 'string' ? 'left' : 'right',
          cursor: 'text', minHeight: 24, background: '#fdf8e8',
        }}
      >
        {fmtDisplay(value, dataType) || '\u00A0'}
      </div>
    )
  }

  const commit = () => { setEditing(false); onStopEdit?.(); if (local !== value) onChange(local) }

  const moveToNext = (e: React.KeyboardEvent, reverse: boolean) => {
    e.preventDefault()
    commit()
    // Find next editable td and click it
    const td = (e.target as HTMLElement).closest('td')
    if (!td) return
    const row = td.closest('tr')
    if (!row) return
    const cells = Array.from(row.querySelectorAll('td'))
    const idx = cells.indexOf(td)
    const step = reverse ? -1 : 1
    // Try same row first
    for (let i = idx + step; i >= 0 && i < cells.length; i += step) {
      const div = cells[i].querySelector('[data-editable-cell]') as HTMLElement
      if (div) { setTimeout(() => div.click(), 0); return }
    }
    // Try next/prev row
    const allRows = Array.from(row.closest('tbody')?.querySelectorAll('tr') || [])
    const rowIdx = allRows.indexOf(row)
    for (let ri = rowIdx + step; ri >= 0 && ri < allRows.length; ri += step) {
      const tds = allRows[ri].querySelectorAll('td')
      const startIdx = reverse ? tds.length - 1 : 1 // skip first (label) column
      for (let i = startIdx; i >= 0 && i < tds.length; i += step) {
        const div = tds[i].querySelector('[data-editable-cell]') as HTMLElement
        if (div) { setTimeout(() => div.click(), 0); return }
      }
    }
  }

  return (
    <input
      ref={inputRef}
      value={local}
      onChange={e => setLocal(e.target.value)}
      onBlur={() => commit()}
      onKeyDown={e => {
        if (e.key === 'Tab') { moveToNext(e, e.shiftKey) }
        else if (e.key === 'Enter') { commit() }
        else if (e.key === 'Escape') { setLocal(value); setEditing(false) }
      }}
      style={{
        width: '100%', border: 'none', outline: 'none', padding: '4px 6px',
        fontSize: 13, textAlign: dataType === 'string' ? 'left' : 'right',
        background: '#e3f2fd', boxSizing: 'border-box',
      }}
    />
  )
}

// ─── Settings dialog (reorder) ───
function SettingsDialog({ open, onClose, order, onReorder, names }: {
  open: boolean; onClose: () => void; order: string[]; onReorder: (ids: string[]) => void; names: Record<string, string>
}) {
  const dragIdx = useRef<number | null>(null)
  const [items, setItems] = useState(order)
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null)
  useEffect(() => { setItems(order) }, [order])
  return (
    <Dialog open={open} onClose={onClose} maxWidth="xs" fullWidth>
      <DialogTitle>Порядок аналитик</DialogTitle>
      <DialogContent>
        <Typography variant="caption" color="textSecondary" sx={{ mb: 1, display: 'block' }}>
          Первая = столбцы, остальные = строки (вложенность по порядку)
        </Typography>
        <List dense>
          {items.map((id, i) => (
            <ListItem key={id} draggable
              onDragStart={() => { dragIdx.current = i }}
              onDragOver={e => { e.preventDefault(); setDragOverIdx(i) }}
              onDrop={() => {
                const from = dragIdx.current
                if (from !== null && from !== i) { const n = [...items]; const [m] = n.splice(from, 1); n.splice(i, 0, m); setItems(n); onReorder(n) }
                dragIdx.current = null; setDragOverIdx(null)
              }}
              onDragEnd={() => { dragIdx.current = null; setDragOverIdx(null) }}
              sx={{ cursor: 'grab', borderTop: dragOverIdx === i ? '2px solid #1976d2' : '2px solid transparent' }}>
              <ListItemIcon sx={{ minWidth: 28 }}><DragIndicatorOutlined sx={{ fontSize: 16, color: '#bbb' }} /></ListItemIcon>
              <ListItemText primary={`${i + 1}. ${names[id] || id}`} />
            </ListItem>
          ))}
        </List>
      </DialogContent>
    </Dialog>
  )
}

// ─── Record picker ───
function RecordPicker({ anchorEl, tree, onSelect, onClose }: {
  anchorEl: HTMLElement | null; tree: RecordNode[]; onSelect: (id: string) => void; onClose: () => void
}) {
  const flat = flattenWithLevel(tree)
  return (
    <Popover open={!!anchorEl} anchorEl={anchorEl} onClose={onClose} anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}>
      <Box sx={{ maxHeight: 300, overflow: 'auto', minWidth: 200, py: 0.5 }}>
        {flat.map(({ node, level }) => (
          <Box key={node.record.id} onClick={() => { onSelect(node.record.id); onClose() }}
            sx={{ px: 2, py: 0.5, pl: 2 + level * 2, cursor: 'pointer', fontSize: 13, '&:hover': { bgcolor: '#f0f0f0' } }}>
            {node.data.name || node.record.id.slice(0, 8)}
          </Box>
        ))}
      </Box>
    </Popover>
  )
}

// ─── Cell rule types ───
type CellRule = 'manual' | 'sum_children' | 'formula'

// ─── Main PivotGrid ───
interface Props {
  sheetId: string; modelId: string
  currentUserId?: string
  mode?: 'data' | 'settings'  // data = view/edit values, settings = cell rules/formulas
}

export default function PivotGrid({ sheetId, modelId, currentUserId, mode: externalMode }: Props) {
  const [bindings, setBindings] = useState<SheetAnalytic[]>([])
  const [analyticsMap, setAnalyticsMap] = useState<Record<string, Analytic>>({})
  const [recordsByAnalytic, setRecordsByAnalytic] = useState<Record<string, RecordNode[]>>({})
  const [cells, setCells] = useState<Record<string, string>>({})
  const [loading, setLoading] = useState(true)
  const [sheetName, setSheetName] = useState('')
  const [modelName, setModelName] = useState('')
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [order, setOrder] = useState<string[]>([])
  const [canEdit, setCanEdit] = useState(true)
  const [collapsedRows, setCollapsedRows] = useState<Set<string>>(new Set())
  const [pinned, setPinned] = useState<Record<string, string>>({})
  const [pickerAnchor, setPickerAnchor] = useState<HTMLElement | null>(null)
  const [pickerAnalyticId, setPickerAnalyticId] = useState<string | null>(null)
  const mode = externalMode || 'data'
  const [cellRules, setCellRules] = useState<Record<string, CellRule>>({})
  const [formulas, setFormulas] = useState<Record<string, string>>({})
  const [formulaEditorOpen, setFormulaEditorOpen] = useState(false)
  const [formulaEditorKey, setFormulaEditorKey] = useState('')
  const [colLevelToggles, setColLevelToggles] = useState<Record<number, boolean>>({})
  // Focus state: [rowIndex, colIndex] in the data grid — always has a focused cell
  const [focusCell, setFocusCell] = useState<[number, number]>([0, 0])
  const [selAnchor, setSelAnchor] = useState<[number, number] | null>(null) // selection anchor for shift+arrows
  const [editingCell, setEditingCell] = useState(false)
  const gridRef = useRef<HTMLTableElement>(null)
  const [colWidths, setColWidths] = useState<Record<number, number>>({})
  const resizingCol = useRef<{ idx: number; startX: number; startW: number } | null>(null)
  const gridBoxRef = useRef<HTMLDivElement>(null)
  // Auto-focus grid on mount
  useEffect(() => { if (!loading) gridBoxRef.current?.focus() }, [loading])

  const load = useCallback(async () => {
    setLoading(true)
    // Load model and sheet names
    const [tree, sa] = await Promise.all([api.getModelTree(modelId), api.listSheetAnalytics(sheetId)])
    setModelName(tree.name || '')
    const sh = (tree.sheets || []).find((s: any) => s.id === sheetId)
    setSheetName(sh?.name || '')
    setBindings(sa)
    const aMap: Record<string, Analytic> = {}
    const rMap: Record<string, RecordNode[]> = {}
    for (const b of sa) {
      const [analytic, recs] = await Promise.all([api.getAnalytic(b.analytic_id), api.listRecords(b.analytic_id)])
      aMap[b.analytic_id] = analytic; rMap[b.analytic_id] = buildRecordTree(recs)
    }
    setAnalyticsMap(aMap); setRecordsByAnalytic(rMap)

    // Load saved view settings
    const defaultOrder = sa.map(b => b.analytic_id)
    try {
      const vs = await api.getViewSettings(sheetId)
      if (vs.order && vs.order.length > 0) {
        // Validate order IDs still exist
        const validIds = new Set(defaultOrder)
        const savedOrder = (vs.order as string[]).filter(id => validIds.has(id))
        // Add any new analytics not in saved order
        for (const id of defaultOrder) if (!savedOrder.includes(id)) savedOrder.push(id)
        setOrder(savedOrder)
      } else {
        setOrder(defaultOrder)
      }
      if (vs.colLevelToggles) setColLevelToggles(vs.colLevelToggles)
      if (vs.pinned) setPinned(vs.pinned)
    } catch {
      setOrder(defaultOrder)
    }

    const cellData = await api.getCells(sheetId, currentUserId)
    const cellMap: Record<string, string> = {}
    const ruleMap: Record<string, CellRule> = {}
    const formulaMap: Record<string, string> = {}
    for (const c of cellData) {
      cellMap[c.coord_key] = c.value ?? ''
      if (c.rule && c.rule !== 'manual') ruleMap[c.coord_key] = c.rule as CellRule
      if (c.formula) formulaMap[c.coord_key] = c.formula
    }
    setCells(cellMap); setCellRules(ruleMap); setFormulas(formulaMap); setLoading(false)
  }, [sheetId])

  // Light reload: only cell values (no structure rebuild, no scroll reset)
  const reloadCells = useCallback(async () => {
    const cellData = await api.getCells(sheetId, currentUserId)
    const cellMap: Record<string, string> = {}
    for (const c of cellData) cellMap[c.coord_key] = c.value ?? ''
    setCells(cellMap)
  }, [sheetId, currentUserId])

  useEffect(() => { load() }, [load])

  // Check edit permission for current sheet
  useEffect(() => {
    if (!currentUserId) { setCanEdit(true); return }
    api.getSheetPermissions(sheetId).then(perms => {
      const myPerm = perms.find((p: any) => p.user_id === currentUserId)
      setCanEdit(myPerm ? !!myPerm.can_edit : true)
    })
  }, [currentUserId, sheetId])

  // Auto-pin analytics where user has access to only 1 record
  useEffect(() => {
    if (!currentUserId) return
    api.getAllowedRecords(currentUserId, sheetId).then(allowed => {
      if (!allowed || Object.keys(allowed).length === 0) return
      setPinned(prev => {
        const next = { ...prev }
        for (const [aId, recordIds] of Object.entries(allowed)) {
          if (recordIds.length === 1) {
            next[aId] = recordIds[0]
          }
        }
        return next
      })
    })
  }, [currentUserId, sheetId])

  // Auto-scroll to keep focused cell visible
  useEffect(() => {
    const box = gridBoxRef.current
    const table = gridRef.current
    if (!box || !table) return
    const [fr, fc] = focusCell
    // Find the cell element: row fr+headerRowCount, col fc+1 (first col is row label)
    const headerRowCount = table.tHead?.rows.length || 1
    const row = table.rows[fr + headerRowCount]
    if (!row) return
    const cell = row.cells[fc + 1] as HTMLTableCellElement | undefined
    if (!cell) return
    // Get sticky column width (first col)
    const stickyWidth = row.cells[0]?.getBoundingClientRect().width || 200
    const cellRect = cell.getBoundingClientRect()
    const boxRect = box.getBoundingClientRect()
    // Horizontal scroll: don't let cell hide under sticky column
    if (cellRect.left < boxRect.left + stickyWidth) {
      box.scrollLeft -= (boxRect.left + stickyWidth - cellRect.left + 4)
    } else if (cellRect.right > boxRect.right) {
      box.scrollLeft += (cellRect.right - boxRect.right + 4)
    }
    // Vertical scroll
    const headerHeight = (table.tHead?.getBoundingClientRect().height || 30)
    if (cellRect.top < boxRect.top + headerHeight) {
      box.scrollTop -= (boxRect.top + headerHeight - cellRect.top + 4)
    } else if (cellRect.bottom > boxRect.bottom) {
      box.scrollTop += (cellRect.bottom - boxRect.bottom + 4)
    }
  }, [focusCell])

  // Column resize handlers
  const handleColResizeStart = (ci: number, e: React.MouseEvent) => {
    e.preventDefault()
    e.stopPropagation()
    const th = (e.target as HTMLElement).closest('th')
    const startW = th?.getBoundingClientRect().width || 90
    resizingCol.current = { idx: ci, startX: e.clientX, startW }
    const onMove = (ev: MouseEvent) => {
      if (!resizingCol.current) return
      const diff = ev.clientX - resizingCol.current.startX
      const newW = Math.max(50, resizingCol.current.startW + diff)
      setColWidths(prev => ({ ...prev, [resizingCol.current!.idx]: newW }))
    }
    const onUp = () => { resizingCol.current = null; document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp) }
    document.addEventListener('mousemove', onMove)
    document.addEventListener('mouseup', onUp)
  }

  // Auto-save view settings on changes
  const saveSettingsTimer = useRef<ReturnType<typeof setTimeout>>()
  useEffect(() => {
    if (loading) return
    clearTimeout(saveSettingsTimer.current)
    saveSettingsTimer.current = setTimeout(() => {
      api.saveViewSettings(sheetId, { order, colLevelToggles, pinned })
    }, 500)
  }, [order, colLevelToggles, pinned, sheetId, loading])

  const analyticNames = useMemo(() => {
    const m: Record<string, string> = {}; for (const [id, a] of Object.entries(analyticsMap)) m[id] = a.name; return m
  }, [analyticsMap])

  const dataType = useMemo(() => {
    if (order.length === 0) return 'sum'
    return analyticsMap[order[0]]?.data_type || 'sum'
  }, [order, analyticsMap])

  const isNumeric = dataType !== 'string'
  const colAnalyticId = order[0]
  const pinnedEntries = order.slice(1).filter(id => !!pinned[id])

  // Pinned leaf analytics are fully fixed — remove from rows
  // Pinned group analytics stay in rows but with filtered tree (pinned node + descendants)
  const pinnedGroupIds = useMemo(() => {
    const s = new Set<string>()
    for (const aId of pinnedEntries) {
      const tree = recordsByAnalytic[aId] || []
      const node = findNodeById(tree, pinned[aId])
      if (node && node.children.length > 0) s.add(aId)
    }
    return s
  }, [pinnedEntries, pinned, recordsByAnalytic])

  const rowAnalyticIds = order.slice(1).filter(id => !pinned[id] || pinnedGroupIds.has(id))

  // Build filtered record trees for pinned group analytics
  const filteredRecordsByAnalytic = useMemo(() => {
    const result: Record<string, RecordNode[]> = { ...recordsByAnalytic }
    for (const aId of pinnedGroupIds) {
      const tree = recordsByAnalytic[aId] || []
      const node = findNodeById(tree, pinned[aId])
      if (node) result[aId] = [node] // Only show this node and its children
    }
    return result
  }, [recordsByAnalytic, pinnedGroupIds, pinned])

  const hasPinnedGroup = pinnedGroupIds.size > 0
  const colTree = recordsByAnalytic[colAnalyticId] || []
  const colDepth = maxDepth(colTree)

  // Column level names for toggles (e.g. level 0 = "Годы", level 1 = "Кварталы")
  const colLevelNames = useMemo(() => {
    if (colDepth <= 1) return []
    const names: string[] = []
    // Walk tree to find names at each non-leaf level
    const walk = (nodes: RecordNode[], lvl: number) => {
      if (nodes.length === 0) return
      if (nodes[0].children.length > 0) {
        if (!names[lvl]) names[lvl] = nodes[0].data.name?.replace(/\s*\d{4}$/, '').replace(/^\d+.*$/, 'Уровень ' + (lvl + 1)) || `Уровень ${lvl + 1}`
        walk(nodes[0].children, lvl + 1)
      }
    }
    walk(colTree, 0)
    // Heuristic: detect level names from first nodes
    const result: { level: number; label: string }[] = []
    const walkForLabel = (nodes: RecordNode[], lvl: number) => {
      if (nodes.length === 0 || nodes[0].children.length === 0) return
      const firstName = nodes[0].data.name || ''
      let label = `Уровень ${lvl + 1}`
      if (/^\d{4}$/.test(firstName)) label = 'Годы'
      else if (/квартал/i.test(firstName)) label = 'Кварталы'
      result.push({ level: lvl, label })
      walkForLabel(nodes[0].children, lvl + 1)
    }
    walkForLabel(colTree, 0)
    return result
  }, [colTree, colDepth])

  // Build display columns: leaf columns + optional group sum columns
  interface DisplayCol {
    node: RecordNode
    isSum: boolean // if true, this is a sum column for a group node
    leafIds: string[] // leaf record IDs to sum (for sum columns)
  }

  const displayCols = useMemo(() => {
    const result: DisplayCol[] = []
    const walkCols = (nodes: RecordNode[], level: number) => {
      for (const n of nodes) {
        if (n.children.length === 0) {
          // Leaf
          result.push({ node: n, isSum: false, leafIds: [n.record.id] })
        } else {
          // Has children - recurse
          walkCols(n.children, level + 1)
          // Add sum column if this level's toggle is on
          if (colLevelToggles[level]) {
            const leaves = getLeaves([n])
            result.push({ node: n, isSum: true, leafIds: leaves.map(l => l.record.id) })
          }
        }
      }
    }
    walkCols(colTree, 0)
    return result
  }, [colTree, colLevelToggles])

  // Build header rows for the display columns
  const headerRows = useMemo(() => {
    if (displayCols.length === 0) return []
    // Simple approach: one header row with column names
    // For multi-level with toggles, build proper nested headers
    if (colDepth <= 1 || colLevelNames.every(l => !colLevelToggles[l.level])) {
      // No group columns enabled or single level - use standard headers
      return buildHeaderRows(colTree)
    }
    // With group columns: flat header showing each displayCol name
    // We still try to build nested headers but include sum cols
    // For simplicity, use a single header row when toggles are active
    return [] // will use displayCols directly for header
  }, [colTree, colDepth, colLevelNames, colLevelToggles, displayCols])

  // Fallback: if headerRows is empty, render a simple header from displayCols
  const useSimpleHeader = headerRows.length === 0 && displayCols.length > 0

  // ─── Build rows ───
  interface RowEntry {
    recordIds: Record<string, string>; label: string; indent: number
    isGroup: boolean; analyticId: string; unit?: string; dragInfo?: { analyticId: string; recordId: string }
    /** The record ID of THIS row's own record in its analytic (for collapse tracking) */
    ownRecordId?: string
    /** Chain of parent record IDs (for hiding when ancestor is collapsed) */
    ancestorRecordIds: string[]
    /** Whether this row has children that can be collapsed */
    hasChildren: boolean
  }
  const rows = useMemo(() => {
    const result: RowEntry[] = []
    const totalAnalytics = rowAnalyticIds.length

    const buildLevel = (ai: number, parentIds: Record<string, string>, baseIndent: number, ancestors: string[]) => {
      if (ai >= totalAnalytics) return
      const aId = rowAnalyticIds[ai]
      for (const { node, level, parentChain } of flattenWithLevelAndParents(filteredRecordsByAnalytic[aId] || [])) {
        const ids = { ...parentIds, [aId]: node.record.id }
        const hasChildren = node.children.length > 0
        const isLastAnalytic = ai === totalAnalytics - 1
        const isGroup = hasChildren || !isLastAnalytic || hasPinnedGroup
        const indent = baseIndent + level
        const allAncestors = [...ancestors, ...parentChain]

        result.push({
          recordIds: ids, label: node.data.name || '', indent,
          isGroup, analyticId: aId, unit: node.data.unit,
          dragInfo: { analyticId: aId, recordId: node.record.id },
          ownRecordId: node.record.id,
          ancestorRecordIds: allAncestors,
          hasChildren,
        })

        if (!hasChildren && !isLastAnalytic) {
          buildLevel(ai + 1, ids, indent + 1, [...allAncestors, node.record.id])
        }
      }
    }

    if (totalAnalytics > 0) buildLevel(0, {}, 0, [])
    else result.push({ recordIds: {}, label: '', indent: 0, isGroup: false, analyticId: '', ancestorRecordIds: [], hasChildren: false })
    return result
  }, [rowAnalyticIds, filteredRecordsByAnalytic, hasPinnedGroup])

  // Filter rows by collapse state
  const visibleRows = useMemo(() => {
    if (collapsedRows.size === 0) return rows
    return rows.filter(row =>
      !row.ancestorRecordIds.some(aid => collapsedRows.has(aid))
    )
  }, [rows, collapsedRows])

  const toggleRowCollapse = useCallback((recordId: string) => {
    setCollapsedRows(prev => {
      const next = new Set(prev)
      if (next.has(recordId)) next.delete(recordId); else next.add(recordId)
      return next
    })
  }, [])

  // ─── Coord key ───
  const makeCoordKey = (rowIds: Record<string, string>, colId: string) => {
    const parts: string[] = []
    for (const aId of order) {
      if (aId === colAnalyticId) parts.push(colId)
      else if (pinned[aId]) parts.push(pinned[aId])
      else if (rowIds[aId]) parts.push(rowIds[aId])
    }
    return parts.join('|')
  }

  // ─── Aggregation ───
  // Expand a row's record IDs to all leaf combinations across all row analytics
  const getAllLeafRecordCombinations = useCallback((row: RowEntry): Record<string, string>[] => {
    let combos: Record<string, string>[] = [{ ...row.recordIds }]

    // Expand row analytics
    for (const aId of rowAnalyticIds) {
      const tree = recordsByAnalytic[aId] || []
      if (!row.recordIds[aId]) {
        const allLeaves = getLeaves(tree)
        const nc: Record<string, string>[] = []
        for (const c of combos) for (const l of allLeaves) nc.push({ ...c, [aId]: l.record.id })
        combos = nc
      } else {
        const node = findNodeById(tree, row.recordIds[aId])
        if (node && node.children.length > 0) {
          const leaves = getLeaves([node])
          const nc: Record<string, string>[] = []
          for (const c of combos) for (const l of leaves) nc.push({ ...c, [aId]: l.record.id })
          combos = nc
        }
      }
    }
    // Also expand pinned analytics that are groups
    for (const aId of pinnedEntries) {
      const tree = recordsByAnalytic[aId] || []
      const node = findNodeById(tree, pinned[aId])
      if (node && node.children.length > 0) {
        const leaves = getLeaves([node])
        const nc: Record<string, string>[] = []
        for (const c of combos) for (const l of leaves) nc.push({ ...c, [aId]: l.record.id })
        combos = nc
      } else if (node) {
        for (const c of combos) c[aId] = node.record.id
      }
    }
    return combos
  }, [rowAnalyticIds, pinnedEntries, pinned, recordsByAnalytic])

  // Coord key using fully expanded combo (not relying on pinned map)
  const makeLeafCoordKey = useCallback((comboIds: Record<string, string>, colId: string) => {
    const parts: string[] = []
    for (const aId of order) {
      if (aId === colAnalyticId) parts.push(colId)
      else if (comboIds[aId]) parts.push(comboIds[aId])
      else if (pinned[aId]) parts.push(pinned[aId])
    }
    return parts.join('|')
  }, [order, colAnalyticId, pinned])

  const computeSum = useCallback((row: RowEntry, colId: string): number | null => {
    const colNode = findNodeById(colTree, colId)
    const colIds = colNode && colNode.children.length > 0 ? getLeaves([colNode]).map(l => l.record.id) : [colId]
    const rowCombos = getAllLeafRecordCombinations(row)
    let sum = 0; let has = false
    for (const combo of rowCombos) {
      for (const cId of colIds) {
        const k = makeLeafCoordKey(combo, cId)
        const v = cells[k]; if (v !== undefined && v !== '') { const n = parseFloat(v); if (!isNaN(n)) { sum += n; has = true } }
      }
    }
    return has ? sum : null
  }, [cells, colTree, getAllLeafRecordCombinations, makeLeafCoordKey])

  // ─── Cell rule resolution ───
  const resolveRule = (coordKey: string, isGroupRow: boolean): CellRule => {
    const explicit = cellRules[coordKey]
    if (explicit) return explicit
    // Default: terminal = manual, non-terminal = sum_children
    return isGroupRow ? 'sum_children' : 'manual'
  }

  // ─── Formula evaluation ───
  // Build name->id maps for resolution
  const recordNameToId = useMemo(() => {
    const map: Record<string, Record<string, string>> = {} // analyticId -> {name -> recordId}
    for (const [aId, tree] of Object.entries(recordsByAnalytic)) {
      const m: Record<string, string> = {}
      const walk = (nodes: RecordNode[]) => {
        for (const n of nodes) { m[n.data.name || ''] = n.record.id; walk(n.children) }
      }
      walk(tree); map[aId] = m
    }
    return map
  }, [recordsByAnalytic])

  const analyticNameToId = useMemo(() => {
    const m: Record<string, string> = {}
    for (const [id, a] of Object.entries(analyticsMap)) m[a.name] = id
    return m
  }, [analyticsMap])

  // Formula evaluation is fully server-side — no client-side eval needed

  // ─── Selection helpers ───
  const selRange = useMemo(() => {
    const [r, c] = focusCell
    if (!selAnchor) return { r1: r, c1: c, r2: r, c2: c }
    const [ar, ac] = selAnchor
    return { r1: Math.min(r, ar), c1: Math.min(c, ac), r2: Math.max(r, ar), c2: Math.max(c, ac) }
  }, [focusCell, selAnchor])
  const isSelected = (ri: number, ci: number) => ri >= selRange.r1 && ri <= selRange.r2 && ci >= selRange.c1 && ci <= selRange.c2

  // ─── Copy / Paste / Delete ───
  const handleCopy = useCallback((e: React.ClipboardEvent | ClipboardEvent) => {
    if (editingCell) return // let native input handle it
    e.preventDefault()
    const { r1, c1, r2, c2 } = selRange
    const lines: string[] = []
    for (let ri = r1; ri <= r2; ri++) {
      const vals: string[] = []
      for (let ci = c1; ci <= c2; ci++) {
        const row = visibleRows[ri]; const col = displayCols[ci]
        if (!row || !col) { vals.push(''); continue }
        if (col.isSum) {
          // sum column — compute aggregated value
          const rowCombos = getAllLeafRecordCombinations(row)
          let sum = 0; let has = false
          for (const combo of rowCombos)
            for (const leafId of col.leafIds) {
              const k = makeLeafCoordKey(combo, leafId)
              const v = cells[k]; if (v !== undefined && v !== '') { const n = parseFloat(v); if (!isNaN(n)) { sum += n; has = true } }
            }
          vals.push(has ? String(sum) : '')
        } else {
          const coordKey = makeCoordKey(row.recordIds, col.node.record.id)
          const rule = resolveRule(coordKey, row.isGroup)
          if (rule === 'sum_children' && isNumeric) {
            const agg = computeSum(row, col.node.record.id)
            vals.push(agg !== null ? String(agg) : '')
          } else if (rule === 'formula') {
            const serverVal = cells[coordKey] ?? ''
            vals.push(serverVal !== '' ? serverVal : '')
          } else {
            vals.push(cells[coordKey] ?? '')
          }
        }
      }
      lines.push(vals.join('\t'))
    }
    const text = lines.join('\n')
    const cb = 'clipboardData' in e ? (e as ClipboardEvent).clipboardData : (e as React.ClipboardEvent).clipboardData
    cb?.setData('text/plain', text)
  }, [editingCell, selRange, visibleRows, displayCols, cells, formulas, isNumeric, computeSum, getAllLeafRecordCombinations, makeLeafCoordKey, makeCoordKey])

  const handlePaste = useCallback(async (e: React.ClipboardEvent | ClipboardEvent) => {
    if (editingCell) return // let native input handle it
    e.preventDefault()
    const cb = 'clipboardData' in e ? (e as ClipboardEvent).clipboardData : (e as React.ClipboardEvent).clipboardData
    const text = cb?.getData('text/plain') || ''
    if (!text) return
    const pasteRows = text.split(/\r?\n/).filter(l => l.length > 0).map(l => l.split('\t'))
    const [startR, startC] = focusCell
    const toSave: { coord_key: string; value: string; data_type: string; user_id?: string }[] = []
    const newCells = { ...cells }
    for (let dr = 0; dr < pasteRows.length; dr++) {
      for (let dc = 0; dc < pasteRows[dr].length; dc++) {
        const ri = startR + dr; const ci = startC + dc
        if (ri >= visibleRows.length || ci >= displayCols.length) continue
        const row = visibleRows[ri]; const col = displayCols[ci]
        if (col.isSum) continue
        const coordKey = makeCoordKey(row.recordIds, col.node.record.id)
        const rule = resolveRule(coordKey, row.isGroup)
        if (rule !== 'manual') continue
        const val = pasteRows[dr][dc]
        newCells[coordKey] = val
        toSave.push({ coord_key: coordKey, value: val, data_type: dataType, user_id: currentUserId })
      }
    }
    setCells(newCells)
    if (toSave.length > 0) { await api.saveCells(sheetId, toSave); reloadCells() }
    // Expand selection to pasted area
    const endR = Math.min(startR + pasteRows.length - 1, rows.length - 1)
    const endC = Math.min(startC + (Math.max(...pasteRows.map(r => r.length)) || 1) - 1, displayCols.length - 1)
    setSelAnchor([endR, endC])
  }, [editingCell, focusCell, visibleRows, displayCols, cells, makeCoordKey, dataType, currentUserId, sheetId])

  const handleDelete = useCallback(async () => {
    if (editingCell) return
    const { r1, c1, r2, c2 } = selRange
    const toSave: { coord_key: string; value: string; data_type: string; user_id?: string }[] = []
    const newCells = { ...cells }
    for (let ri = r1; ri <= r2; ri++) {
      for (let ci = c1; ci <= c2; ci++) {
        const row = visibleRows[ri]; const col = displayCols[ci]
        if (!row || !col || col.isSum) continue
        const coordKey = makeCoordKey(row.recordIds, col.node.record.id)
        const rule = resolveRule(coordKey, row.isGroup)
        if (rule !== 'manual') continue
        newCells[coordKey] = ''
        toSave.push({ coord_key: coordKey, value: '', data_type: dataType, user_id: currentUserId })
      }
    }
    setCells(newCells)
    if (toSave.length > 0) { await api.saveCells(sheetId, toSave); reloadCells() }
  }, [editingCell, selRange, visibleRows, displayCols, cells, makeCoordKey, dataType, currentUserId, sheetId, reloadCells])

  // ─── Context menu + history ───
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; coordKey: string } | null>(null)
  const [historyOpen, setHistoryOpen] = useState(false)
  const [historyKey, setHistoryKey] = useState('')
  const [historyData, setHistoryData] = useState<any[]>([])

  const handleContextMenu = (e: React.MouseEvent, coordKey: string, rule: CellRule) => {
    if (rule !== 'manual') return
    e.preventDefault()
    setCtxMenu({ x: e.clientX, y: e.clientY, coordKey })
  }

  const showHistory = async (coordKey: string) => {
    setCtxMenu(null)
    setHistoryKey(coordKey)
    const data = await api.getCellHistory(sheetId, coordKey)
    setHistoryData(data)
    setHistoryOpen(true)
  }

  // ─── Save (auto-recalc on backend, then reload) ───
  const handleCellSave = async (coordKey: string, value: string) => {
    setCells(prev => ({ ...prev, [coordKey]: value }))
    await api.saveCells(sheetId, [{ coord_key: coordKey, value, data_type: dataType, user_id: currentUserId }])
    reloadCells() // reload only values, keep scroll/focus
  }

  const handleReorder = (newOrder: string[]) => {
    setOrder(newOrder)
    if (pinned[newOrder[0]]) setPinned(prev => { const n = { ...prev }; delete n[newOrder[0]]; return n })
  }
  const handlePin = (aId: string, rId: string) => setPinned(prev => ({ ...prev, [aId]: rId }))
  const handleUnpin = (aId: string) => setPinned(prev => { const n = { ...prev }; delete n[aId]; return n })

  if (loading) return (
    <Box sx={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', bgcolor: '#fafafa' }}>
      <Typography color="textSecondary">Загрузка...</Typography>
    </Box>
  )

  const getIcon = (aId: string) => {
    const nm = analyticsMap[aId]?.icon; if (!nm) return null
    const I = (Icons as any)[nm]; return I ? <I sx={{ fontSize: 14 }} /> : null
  }

  return (
    <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', bgcolor: '#fff' }}>
      {/* ─── Grid toolbar (compact) ─── */}
      <Box
        sx={{ display: 'flex', alignItems: 'center', px: 1, py: 0.5, borderBottom: '1px solid #f0f0f0', gap: 1, flexWrap: 'wrap', minHeight: 32 }}
        onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy' }}
        onDrop={e => {
          e.preventDefault()
          try {
            const data = JSON.parse(e.dataTransfer.getData('text/plain'))
            if (data.analyticId && data.recordId) handlePin(data.analyticId, data.recordId)
          } catch {}
        }}
      >
        <Tooltip title="Порядок аналитик">
          <IconButton size="small" onClick={() => setSettingsOpen(true)}>
            <FormatListNumberedOutlined fontSize="small" />
          </IconButton>
        </Tooltip>

        <Tooltip title="Выгрузить в Excel">
          <IconButton size="small" onClick={() => window.open(api.exportSheetExcelUrl(sheetId), '_blank')}>
            <FileDownloadOutlined fontSize="small" />
          </IconButton>
        </Tooltip>

        <Tooltip title="Загрузить из Excel">
          <IconButton size="small" component="label">
            <FileUploadOutlined fontSize="small" />
            <input type="file" hidden accept=".xlsx" onChange={async e => {
              const f = e.target.files?.[0]
              if (!f) return
              await api.importSheetExcel(sheetId, f)
              e.target.value = ''
              load()
            }} />
          </IconButton>
        </Tooltip>

        <Tooltip title="Рассчитать формулы">
          <IconButton size="small" onClick={async () => {
            await api.calculateSheet(sheetId)
            load()
          }}>
            <CalculateOutlined fontSize="small" />
          </IconButton>
        </Tooltip>

        {/* Column level toggles */}
        {colLevelNames.length > 0 && colLevelNames.map(({ level, label }) => (
          <Chip
            key={`lvl-${level}`}
            size="small"
            label={`Σ ${label}`}
            color={colLevelToggles[level] ? 'primary' : 'default'}
            variant={colLevelToggles[level] ? 'filled' : 'outlined'}
            onClick={() => setColLevelToggles(prev => ({ ...prev, [level]: !prev[level] }))}
            sx={{ fontSize: 11 }}
          />
        ))}

        {/* Pinned chips */}
        {pinnedEntries.map(aId => {
          const tree = recordsByAnalytic[aId] || []
          const node = findNodeById(tree, pinned[aId])
          return (
            <Chip key={aId} size="small" icon={getIcon(aId) || undefined}
              label={`${analyticNames[aId]}: ${node?.data.name || '?'}`}
              onClick={e => { setPickerAnchor(e.currentTarget); setPickerAnalyticId(aId) }}
              onDelete={() => handleUnpin(aId)} sx={{ fontSize: 12 }} />
          )
        })}

        <Box sx={{ flex: 1, display: 'flex', justifyContent: 'center', minWidth: 0 }}>
          <Typography noWrap sx={{ fontSize: 13, color: '#555' }}>
            {modelName}{modelName && sheetName ? ' → ' : ''}{sheetName}
          </Typography>
        </Box>

        {!canEdit && <Chip size="small" label="только чтение" sx={{ fontSize: 11, bgcolor: '#fff5f5', color: '#c62828' }} />}
      </Box>

      {/* ─── Grid ─── */}
      <Box ref={gridBoxRef} sx={{ flex: 1, overflow: 'auto', outline: 'none' }}
        tabIndex={0}
        onCopy={handleCopy}
        onPaste={handlePaste}
        onKeyDown={e => {
          if (editingCell || formulaEditorOpen || settingsOpen) return
          const totalRows = visibleRows.length
          const totalCols = displayCols.length
          const [fr, fc] = focusCell
          const shift = e.shiftKey
          const moveWithSelection = (nr: number, nc: number) => {
            if (shift) {
              if (!selAnchor) setSelAnchor([fr, fc])
            } else {
              setSelAnchor(null)
            }
            setFocusCell([nr, nc])
          }
          switch (e.key) {
            case 'ArrowDown': e.preventDefault(); moveWithSelection(Math.min(fr + 1, totalRows - 1), fc); break
            case 'ArrowUp': e.preventDefault(); moveWithSelection(Math.max(fr - 1, 0), fc); break
            case 'ArrowRight': e.preventDefault(); moveWithSelection(fr, Math.min(fc + 1, totalCols - 1)); break
            case 'ArrowLeft': e.preventDefault(); moveWithSelection(fr, Math.max(fc - 1, 0)); break
            case 'Tab': e.preventDefault(); setSelAnchor(null); if (shift) { if (fc > 0) setFocusCell([fr, fc - 1]); else if (fr > 0) setFocusCell([fr - 1, totalCols - 1]) } else { if (fc < totalCols - 1) setFocusCell([fr, fc + 1]); else if (fr < totalRows - 1) setFocusCell([fr + 1, 0]) } break
            case 'Enter': e.preventDefault(); setEditingCell(true); break
            case 'Escape': e.preventDefault(); setEditingCell(false); setSelAnchor(null); break
            case 'Delete': case 'Backspace': e.preventDefault(); handleDelete(); break
            case 'a':
              if (e.ctrlKey || e.metaKey) {
                e.preventDefault()
                setSelAnchor([0, 0])
                setFocusCell([totalRows - 1, totalCols - 1])
              }
              break
            default:
              // Start typing immediately enters edit mode
              if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                setSelAnchor(null)
                setEditingCell(true)
              }
              break
          }
        }}
      >
        <table ref={gridRef} style={{ borderCollapse: 'collapse', fontSize: 13, minWidth: '100%' }}>
          <thead>
            {useSimpleHeader ? (
              <tr>
                <th style={{
                  border: '1px solid #e0e0e0', padding: '4px 8px', background: '#f5f5f5',
                  minWidth: 200, textAlign: 'left', position: 'sticky', left: 0, zIndex: 2, borderRight: '2px solid #bdbdbd',
                }}>
                  {rowAnalyticIds.map(id => analyticNames[id]).join(' / ') || '—'}
                </th>
                {displayCols.map((dc, ci) => (
                  <th key={`${dc.node.record.id}-${dc.isSum ? 's' : 'l'}`} style={{
                    border: '1px solid #e0e0e0', padding: '4px 8px',
                    background: dc.isSum ? '#f5f5f5' : '#f5f5f5',
                    textAlign: 'center', whiteSpace: 'nowrap',
                    width: colWidths[ci] || 90, minWidth: 50,
                    fontWeight: dc.isSum ? 700 : 400,
                    position: 'relative',
                  }}>
                    {dc.isSum ? `Σ ${dc.node.data.name || ''}` : (dc.node.data.name || '')}
                    <span
                      style={{ position: 'absolute', right: 0, top: 0, bottom: 0, width: 4, cursor: 'col-resize' }}
                      onMouseDown={e => handleColResizeStart(ci, e)}
                    />
                  </th>
                ))}
              </tr>
            ) : (
              headerRows.map((row, ri) => (
                <tr key={ri}>
                  {ri === 0 && (
                    <th rowSpan={headerRows.length} style={{
                      border: '1px solid #e0e0e0', padding: '4px 8px', background: '#f5f5f5',
                      minWidth: 200, textAlign: 'left', verticalAlign: 'bottom',
                      position: 'sticky', left: 0, zIndex: 2, borderRight: '2px solid #bdbdbd',
                    }}>
                      {rowAnalyticIds.map(id => analyticNames[id]).join(' / ') || '—'}
                    </th>
                  )}
                  {row.map(({ node, colspan, rowspan }) => (
                    <th key={node.record.id} colSpan={colspan} rowSpan={rowspan} style={{
                      border: '1px solid #e0e0e0', padding: '4px 8px', background: '#f5f5f5',
                      textAlign: 'center', whiteSpace: 'nowrap', minWidth: 90,
                    }}>{node.data.name || ''}</th>
                  ))}
                </tr>
              ))
            )}
          </thead>
          <tbody>
            {visibleRows.map((row, ri) => {
              const rowDt = unitToDataType(row.unit, dataType)
              return (<tr key={ri}>
                <td
                  draggable={!!row.dragInfo}
                  onDragStart={e => {
                    if (row.dragInfo) {
                      e.dataTransfer.setData('text/plain', JSON.stringify(row.dragInfo))
                      e.dataTransfer.effectAllowed = 'copy'
                    }
                  }}
                  style={{
                    border: '1px solid #e0e0e0', padding: '2px 8px', paddingLeft: 8 + row.indent * 16,
                    whiteSpace: 'nowrap', fontWeight: row.isGroup ? 600 : 400,
                    background: row.isGroup ? '#fafafa' : '#fff',
                    position: 'sticky', left: 0, zIndex: 1, borderRight: '2px solid #bdbdbd',
                    cursor: row.dragInfo ? 'grab' : 'default',
                  }}>
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: 2 }}>
                    {row.hasChildren ? (
                      <span
                        style={{ display: 'inline-flex', cursor: 'pointer', opacity: 0.5, marginLeft: -2 }}
                        onClick={e => { e.stopPropagation(); if (row.ownRecordId) toggleRowCollapse(row.ownRecordId) }}
                      >
                        {collapsedRows.has(row.ownRecordId || '') ?
                          <Icons.ChevronRightOutlined sx={{ fontSize: 16 }} /> :
                          <Icons.ExpandMoreOutlined sx={{ fontSize: 16 }} />}
                      </span>
                    ) : (
                      <span style={{ width: 16, display: 'inline-block' }} />
                    )}
                    {row.analyticId && (() => { const ic = getIcon(row.analyticId); return ic ? <span style={{ display: 'inline-flex', opacity: 0.5 }}>{ic}</span> : null })()}
                    {row.label || '—'}
                  </span>
                </td>
                {displayCols.map((col, ci) => {
                  const isFocused = focusCell[0] === ri && focusCell[1] === ci
                  const isSel = isSelected(ri, ci)
                  const focusBorder = isFocused ? '2px solid #1976d2' : isSel ? '1px solid #90caf9' : '1px solid #e0e0e0'
                  const selBg = isSel && !isFocused ? 'rgba(25,118,210,0.08)' : undefined
                  const cellClick = (e: React.MouseEvent) => {
                    if (e.shiftKey) {
                      if (!selAnchor) setSelAnchor([focusCell[0], focusCell[1]])
                      setFocusCell([ri, ci])
                    } else {
                      setSelAnchor(null); setFocusCell([ri, ci])
                    }
                    setEditingCell(false)
                  }

                  // For sum columns: always show aggregated value
                  if (col.isSum) {
                    const rowCombos = getAllLeafRecordCombinations(row)
                    let sum = 0; let has = false
                    for (const combo of rowCombos) {
                      for (const leafId of col.leafIds) {
                        const k = makeLeafCoordKey(combo, leafId)
                        const v = cells[k]
                        if (v !== undefined && v !== '') { const n = parseFloat(v); if (!isNaN(n)) { sum += n; has = true } }
                      }
                    }
                    return (
                      <td key={`${col.node.record.id}-s`} onClick={cellClick} style={{
                        border: focusBorder, padding: '4px 6px',
                        textAlign: 'right', color: '#555', background: selBg || '#fff', fontSize: 13,
                        fontWeight: 600,
                      }}>
                        {has ? fmtDisplay(String(sum), rowDt) : ''}
                      </td>
                    )
                  }

                  const colRecId = col.node.record.id
                  const coordKey = makeCoordKey(row.recordIds, colRecId)
                  const rule = resolveRule(coordKey, row.isGroup)
                  const editable = rule === 'manual'

                  if (mode === 'settings') {
                    const ruleLabel = rule === 'manual' ? 'ввод' : rule === 'sum_children' ? 'сумма' : 'формула'
                    const ruleColor = rule === 'formula' ? '#1565c0' : rule === 'sum_children' ? '#2e7d32' : '#666'
                    const formulaText = formulas[coordKey] || ''
                    const cellContent = rule === 'formula' && formulaText ? formulaText : ruleLabel

                    if (isFocused) {
                      return (
                        <td key={colRecId} onClick={cellClick} style={{ border: focusBorder, padding: 0, background: '#fafbfc' }}>
                          <Box sx={{ display: 'flex', alignItems: 'center' }}>
                            <Select
                              value={rule}
                              variant="standard"
                              disableUnderline
                              onChange={async e => {
                                const newRule = e.target.value as CellRule
                                setCellRules(prev => ({ ...prev, [coordKey]: newRule }))
                                await api.saveCells(sheetId, [{ coord_key: coordKey, rule: newRule }])
                                reloadCells()
                              }}
                              sx={{ fontSize: 11, px: 0.5, minWidth: 70, '& .MuiSelect-select': { py: 0.25 } }}
                            >
                              <MenuItem value="manual" sx={{ fontSize: 12 }}>✎ Ввод</MenuItem>
                              <MenuItem value="sum_children" sx={{ fontSize: 12 }}>Σ Сумма</MenuItem>
                              <MenuItem value="formula" sx={{ fontSize: 12 }}>ƒ Формула</MenuItem>
                            </Select>
                            {rule === 'formula' && (
                              <>
                                <Box sx={{ flex: 1, fontSize: 11, color: '#1565c0', px: 0.5, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {formulaText}
                                </Box>
                                <IconButton size="small" onClick={() => { setFormulaEditorKey(coordKey); setFormulaEditorOpen(true) }}
                                  sx={{ p: 0.25 }}>
                                  <MoreHorizOutlined sx={{ fontSize: 14 }} />
                                </IconButton>
                              </>
                            )}
                          </Box>
                        </td>
                      )
                    }
                    return (
                      <td key={colRecId} onClick={cellClick} style={{
                        border: focusBorder, padding: '2px 4px', background: selBg || '#fafbfc',
                        fontSize: 11, color: ruleColor, cursor: 'pointer',
                        wordBreak: 'break-word', minWidth: 80,
                      }}>
                        {cellContent}
                      </td>
                    )
                  }

                  // Data mode
                  // Formula cell — use server-computed value from cells[]
                  if (rule === 'formula') {
                    const fText = formulas[coordKey] || ''
                    const serverVal = cells[coordKey] ?? ''
                    const num = serverVal !== '' ? parseFloat(serverVal) : null
                    const result = num !== null && !isNaN(num) ? num : null
                    return (
                      <td key={colRecId} onClick={cellClick} style={{
                        border: focusBorder, padding: '4px 6px',
                        textAlign: 'right', color: '#555', background: selBg || '#fff', fontSize: 13,
                      }} title={fText ? `ƒ ${fText}` : 'Формула не задана'}>
                        {result !== null && !isNaN(result) ? fmtDisplay(String(result), rowDt) : ''}
                      </td>
                    )
                  }

                  // Sum cell
                  if (rule === 'sum_children' && isNumeric) {
                    const agg = computeSum(row, colRecId)
                    return (
                      <td key={colRecId} onClick={cellClick} style={{
                        border: focusBorder, padding: '4px 6px',
                        textAlign: 'right', color: '#666', background: selBg || '#fff', fontSize: 13,
                      }}>
                        {agg !== null ? fmtDisplay(String(agg), rowDt) : ''}
                      </td>
                    )
                  }

                  if (rule === 'sum_children') {
                    return (
                      <td key={colRecId} onClick={cellClick} style={{
                        border: focusBorder, padding: '4px 6px',
                        background: selBg || '#fff', color: '#666', fontSize: 13,
                      }}>{cells[coordKey] ?? ''}</td>
                    )
                  }

                  // Manual input cell
                  const shouldEdit = isFocused && editingCell && canEdit
                  const readOnlyBg = !canEdit ? '#fff5f5' : undefined
                  return (
                    <td key={colRecId} onClick={cellClick}
                      style={{ border: focusBorder, padding: 0, background: readOnlyBg }}
                      onContextMenu={e => handleContextMenu(e, coordKey, rule)}>
                      <PivotCell
                        value={cells[coordKey] ?? ''}
                        onChange={val => handleCellSave(coordKey, val)}
                        dataType={rowDt}
                        editable={canEdit}
                        forceEdit={shouldEdit}
                        onStopEdit={() => setEditingCell(false)}
                      />
                    </td>
                  )
                })}
              </tr>)
            })}
          </tbody>
        </table>
      </Box>

      <SettingsDialog open={settingsOpen} onClose={() => setSettingsOpen(false)}
        order={order} onReorder={handleReorder} names={analyticNames} />

      {pickerAnalyticId && (
        <RecordPicker anchorEl={pickerAnchor} tree={recordsByAnalytic[pickerAnalyticId] || []}
          onSelect={id => handlePin(pickerAnalyticId!, id)}
          onClose={() => { setPickerAnchor(null); setPickerAnalyticId(null) }} />
      )}

      <FormulaEditor
        open={formulaEditorOpen}
        formula={formulas[formulaEditorKey] || ''}
        onSave={async text => {
          setFormulas(prev => ({ ...prev, [formulaEditorKey]: text }))
          await api.saveCells(sheetId, [{ coord_key: formulaEditorKey, formula: text, rule: 'formula' }])
          reloadCells()
        }}
        onClose={() => setFormulaEditorOpen(false)}
        modelId={modelId}
      />

      {/* Context menu */}
      {ctxMenu && (
        <Box
          sx={{
            position: 'fixed', left: ctxMenu.x, top: ctxMenu.y, zIndex: 1400,
            bgcolor: '#fff', border: '1px solid #e0e0e0', borderRadius: 1,
            boxShadow: 2, py: 0.5, minWidth: 120,
          }}
          onClick={() => setCtxMenu(null)}
        >
          <Box sx={{ px: 2, py: 0.5, cursor: 'pointer', fontSize: 13, '&:hover': { bgcolor: '#f0f0f0' } }}
            onClick={() => showHistory(ctxMenu.coordKey)}>
            История
          </Box>
        </Box>
      )}
      {ctxMenu && <Box sx={{ position: 'fixed', inset: 0, zIndex: 1399 }} onClick={() => setCtxMenu(null)} />}

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
                {historyData.map((h, i) => (
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
    </Box>
  )
}
