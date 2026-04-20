import { useState, useEffect, useCallback, useRef } from 'react'
import {
  Box, Typography, IconButton, Tooltip, Table, TableHead, TableBody,
  TableRow, TableCell, TextField, MenuItem, Select, FormControl, InputLabel,
} from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import ChevronRightOutlined from '@mui/icons-material/ChevronRightOutlined'
import FileDownloadOutlined from '@mui/icons-material/FileDownloadOutlined'
import FileUploadOutlined from '@mui/icons-material/FileUploadOutlined'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import { usePending } from '../../store/PendingContext'
import * as api from '../../api'
import type { AnalyticField, AnalyticRecord } from '../../types'
import IndicatorFormulasPanel from '../sheet/IndicatorFormulasPanel'
import AnalyticSettings from './AnalyticSettings'
import AnalyticFields from './AnalyticFields'
import Splitter from '../../components/Splitter'

function RecordCellInput({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  const [local, setLocal] = useState(value)
  useEffect(() => { setLocal(value) }, [value])
  return (
    <TextField
      variant="standard"
      fullWidth
      value={local}
      onChange={e => setLocal(e.target.value)}
      onBlur={() => { if (local !== value) onChange(local) }}
      InputProps={{ disableUnderline: true, sx: { fontSize: 13 } }}
    />
  )
}

interface Props {
  analyticId: string
  modelId?: string
  onRefresh?: () => void
}

interface TreeNode {
  record: AnalyticRecord
  data: Record<string, any>
  children: TreeNode[]
  level: number
}

export default function AnalyticRecordsGrid({ analyticId, modelId, onRefresh }: Props) {
  const [fields, setFields] = useState<AnalyticField[]>([])
  const [records, setRecords] = useState<AnalyticRecord[]>([])
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set())
  const fileRef = useRef<HTMLInputElement>(null)
  const { addOp, getOverrides } = usePending()
  // Sheets where this analytic is the main one.
  const [mainSheets, setMainSheets] = useState<{ id: string; name: string }[]>([])
  // Record currently showing its formulas in the right pane.
  const [selectedRecordId, setSelectedRecordId] = useState<string | null>(null)
  // Which of the main sheets the panel targets (when > 1).
  const [activeSheetId, setActiveSheetId] = useState<string | null>(null)
  // Resizable right-pane width.
  const [panelWidth, setPanelWidth] = useState<number>(420)
  // Formulas map: recordId → { leaf, consolidation } for display in the grid.
  const [formulas, setFormulas] = useState<Record<string, { leaf: string; consolidation: string }>>({})

  const load = useCallback(async () => {
    const [fs, rs] = await Promise.all([api.listFields(analyticId), api.listRecords(analyticId)])
    setFields(fs)
    setRecords(rs)
  }, [analyticId])

  useEffect(() => { load() }, [load])

  // Figure out which sheets have this analytic as main.
  useEffect(() => {
    let cancelled = false
    ;(async () => {
      try {
        const a = await api.getAnalytic(analyticId)
        const mid = (a as any).model_id
        if (!mid) return
        const sheets = await api.listSheets(mid)
        const out: { id: string; name: string }[] = []
        for (const s of sheets) {
          const m = await api.getMainAnalytic(s.id)
          if (m.analytic_id === analyticId) out.push({ id: s.id, name: s.name })
        }
        if (!cancelled) {
          setMainSheets(out)
          if (out.length > 0 && !activeSheetId) setActiveSheetId(out[0].id)
        }
      } catch { /* ignore */ }
    })()
    return () => { cancelled = true }
  }, [analyticId])

  // Load formulas for all records in one batch request.
  useEffect(() => {
    if (!activeSheetId || records.length === 0) return
    let cancelled = false
    ;(async () => {
      try {
        const all = await api.getAllIndicatorRules(activeSheetId)
        if (!cancelled) setFormulas(all)
      } catch { /* no rules */ }
    })()
    return () => { cancelled = true }
  }, [activeSheetId, records])

  const buildTree = (): TreeNode[] => {
    const byParent: Record<string, AnalyticRecord[]> = { root: [] }
    for (const r of records) {
      const key = r.parent_id || 'root'
      ;(byParent[key] ||= []).push(r)
    }
    const build = (parentId: string | null, level: number): TreeNode[] => {
      const items = byParent[parentId || 'root'] || []
      return items.map(r => {
        const data = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
        return { record: r, data, children: build(r.id, level + 1), level }
      })
    }
    return build(null, 0)
  }

  const flattenTree = (nodes: TreeNode[]): TreeNode[] => {
    const result: TreeNode[] = []
    const walk = (ns: TreeNode[]) => {
      for (const n of ns) {
        result.push(n)
        if (!collapsed.has(n.record.id)) walk(n.children)
      }
    }
    walk(nodes)
    return result
  }

  const toggle = (id: string) => {
    setCollapsed(prev => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }

  const handleAdd = async (parentId: string | null) => {
    const data: Record<string, any> = {}
    if (fields.length > 0) data[fields[0].code] = ''
    const created = await api.createRecord(analyticId, { parent_id: parentId, sort_order: records.length, data_json: data })
    // Expand parent if adding child
    if (parentId) {
      setCollapsed(prev => {
        const next = new Set(prev)
        next.delete(parentId)
        return next
      })
    }
    // Add locally — no reload needed
    if (created?.id) {
      setRecords(prev => [...prev, { id: created.id, analytic_id: analyticId, parent_id: parentId, sort_order: records.length, data_json: JSON.stringify(data) } as AnalyticRecord])
    }
  }

  const handleDelete = async (id: string) => {
    await api.deleteRecord(analyticId, id)
    // Remove locally (including children)
    setRecords(prev => {
      const toRemove = new Set<string>([id])
      let changed = true
      while (changed) {
        changed = false
        for (const r of prev) {
          if (r.parent_id && toRemove.has(r.parent_id) && !toRemove.has(r.id)) {
            toRemove.add(r.id)
            changed = true
          }
        }
      }
      return prev.filter(r => !toRemove.has(r.id))
    })
  }

  const handleCellChange = (record: AnalyticRecord, fieldCode: string, value: string) => {
    const parsed = typeof record.data_json === 'string' ? JSON.parse(record.data_json) : record.data_json
    const currentData: Record<string, any> = { ...parsed }
    currentData[fieldCode] = value
    // Update local state immutably
    setRecords(prev => prev.map(r =>
      r.id === record.id ? { ...r, data_json: JSON.stringify(currentData) } : r
    ))
    // Add to pending
    const existing = getOverrides(`record:${record.id}`)
    const existingData = existing?.data_json
      ? (typeof existing.data_json === 'string' ? JSON.parse(existing.data_json) : existing.data_json)
      : {}
    const mergedData = { ...existingData, ...currentData }
    addOp({
      key: `record:${record.id}`,
      type: 'updateRecord',
      id: record.id,
      parentId: analyticId,
      data: {
        parent_id: record.parent_id,
        sort_order: record.sort_order,
        data_json: mergedData,
      },
    })
  }

  const handleExport = () => {
    window.open(api.exportExcelUrl(analyticId), '_blank')
  }

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (!file) return
    await api.importExcel(analyticId, file)
    load()
    e.target.value = ''
  }

  const tree = buildTree()
  const flat = flattenTree(tree)

  const selectedRecord = selectedRecordId ? records.find(r => r.id === selectedRecordId) : null
  const selectedRecordName = selectedRecord
    ? (() => {
        try {
          const d = typeof selectedRecord.data_json === 'string'
            ? JSON.parse(selectedRecord.data_json)
            : selectedRecord.data_json
          return (d && (d[fields[0]?.code] || d.name)) || selectedRecord.id.slice(0, 6)
        } catch { return selectedRecord.id.slice(0, 6) }
      })()
    : ''
  const showFormulaPanel = !!selectedRecordId && mainSheets.length > 0 && !!activeSheetId && !!modelId

  return (
    <Box sx={{ display: 'flex', minWidth: 0, width: '100%', height: '100%' }}>
      <Box sx={{ flex: 1, minWidth: 0, overflow: 'auto', p: 1 }}>
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
        <Typography variant="subtitle1">Записи</Typography>
        <Tooltip title="Добавить запись">
          <IconButton size="small" onClick={() => handleAdd(null)}><AddOutlined fontSize="small" /></IconButton>
        </Tooltip>
        <Box sx={{ flex: 1 }} />
        <Tooltip title="Импорт из Excel">
          <IconButton size="small" onClick={() => fileRef.current?.click()}>
            <FileUploadOutlined fontSize="small" />
          </IconButton>
        </Tooltip>
        <Tooltip title="Экспорт в Excel">
          <IconButton size="small" onClick={handleExport}>
            <FileDownloadOutlined fontSize="small" />
          </IconButton>
        </Tooltip>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={handleImport} />
      </Box>

      {fields.length > 0 && (
        <Table size="small" sx={{ '& td, & th': { py: 0.5, px: 1 } }}>
          <TableHead>
            <TableRow>
              {fields.map(f => <TableCell key={f.id}>{f.name}</TableCell>)}
              {mainSheets.length > 0 && (
                <TableCell sx={{ minWidth: 200 }}>Формула</TableCell>
              )}
            </TableRow>
          </TableHead>
          <TableBody>
            {flat.map(node => {
              const hasChildren = node.children.length > 0
              const isCollapsed = collapsed.has(node.record.id)

              const isSelected = selectedRecordId === node.record.id
              return (
                <TableRow
                  key={node.record.id}
                  hover
                  selected={isSelected}
                  onClick={() => { if (mainSheets.length > 0) setSelectedRecordId(node.record.id) }}
                  sx={{
                    cursor: mainSheets.length > 0 ? 'pointer' : 'default',
                    '& .row-actions': { opacity: 0 },
                    '&:hover .row-actions': { opacity: 1 },
                  }}
                >
                  {fields.map((f, fi) => (
                    <TableCell key={f.id} sx={fi === 0 ? { pl: node.level * 3 + 1 } : undefined}>
                      {fi === 0 ? (
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0 }}>
                          <Box className="row-actions" sx={{ display: 'flex', gap: 0, transition: 'opacity 0.15s', flexShrink: 0 }}>
                            <Tooltip title="Добавить дочерний">
                              <IconButton size="small" onClick={() => handleAdd(node.record.id)}>
                                <AddOutlined sx={{ fontSize: 14 }} />
                              </IconButton>
                            </Tooltip>
                            <Tooltip title="Удалить">
                              <IconButton size="small" onClick={() => handleDelete(node.record.id)}>
                                <DeleteOutlineOutlined sx={{ fontSize: 14 }} />
                              </IconButton>
                            </Tooltip>
                          </Box>
                          {hasChildren ? (
                            <IconButton size="small" sx={{ flexShrink: 0 }} onClick={() => toggle(node.record.id)}>
                              {isCollapsed
                                ? <ChevronRightOutlined sx={{ fontSize: 16 }} />
                                : <ExpandMoreOutlined sx={{ fontSize: 16 }} />}
                            </IconButton>
                          ) : <Box sx={{ width: 28, flexShrink: 0 }} />}
                          <RecordCellInput
                            value={node.data[f.code] ?? ''}
                            onChange={val => handleCellChange(node.record, f.code, val)}
                          />
                        </Box>
                      ) : (
                        <RecordCellInput
                          value={node.data[f.code] ?? ''}
                          onChange={val => handleCellChange(node.record, f.code, val)}
                        />
                      )}
                    </TableCell>
                  ))}
                  {mainSheets.length > 0 && (() => {
                    const f = formulas[node.record.id]
                    const txt = f?.leaf || f?.consolidation || ''
                    return (
                      <TableCell
                        data-testid={`formula-cell-${node.record.id}`}
                        onClick={e => { e.stopPropagation(); setSelectedRecordId(node.record.id) }}
                        sx={{
                          cursor: 'pointer',
                          fontFamily: 'monospace',
                          fontSize: 12,
                          color: isSelected ? (txt ? 'text.primary' : 'text.disabled') : 'transparent',
                          maxWidth: 300,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap',
                          '&:hover': { color: txt ? 'text.primary' : 'text.disabled', bgcolor: '#e3f2fd' },
                        }}
                        title={txt || 'Нет формулы — нажмите для настройки'}
                      >
                        {txt || '—'}
                      </TableCell>
                    )
                  })()}
                </TableRow>
              )
            })}
          </TableBody>
        </Table>
      )}

      {records.length === 0 && fields.length > 0 && (
        <Typography variant="body2" color="textSecondary" sx={{ mt: 1 }}>
          Нет записей. Добавьте запись или импортируйте из Excel.
        </Typography>
      )}
      </Box>

      <Splitter onResize={d => setPanelWidth(w => Math.max(260, Math.min(900, w - d)))} />
      <Box sx={{ width: panelWidth, flexShrink: 0, borderLeft: 1, borderColor: 'divider', display: 'flex', flexDirection: 'column', minWidth: 0, overflow: 'auto' }}>
        {showFormulaPanel ? (
          <>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.5, borderBottom: 1, borderColor: 'divider' }}>
              {mainSheets.length > 1 ? (
                <FormControl size="small" sx={{ minWidth: 160 }}>
                  <InputLabel>Лист</InputLabel>
                  <Select
                    label="Лист" value={activeSheetId || ''}
                    onChange={e => setActiveSheetId(String(e.target.value))}
                  >
                    {mainSheets.map(s => (
                      <MenuItem key={s.id} value={s.id}>{s.name}</MenuItem>
                    ))}
                  </Select>
                </FormControl>
              ) : (
                <Typography variant="caption" color="text.secondary" noWrap sx={{ flex: 1, minWidth: 0 }}>
                  {mainSheets[0]?.name}
                </Typography>
              )}
              <Box sx={{ flex: 1 }} />
              <Tooltip title="Назад к настройкам">
                <IconButton size="small" onClick={() => setSelectedRecordId(null)}>
                  <CloseOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
            </Box>
            <Box sx={{ flex: 1, overflow: 'auto' }}>
              <IndicatorFormulasPanel
                sheetId={activeSheetId!}
                modelId={modelId!}
                indicatorId={selectedRecordId!}
                indicatorName={selectedRecordName}
              />
            </Box>
          </>
        ) : (
          <Box sx={{ overflow: 'auto', p: 1 }}>
            {modelId && (
              <AnalyticSettings
                analyticId={analyticId}
                modelId={modelId}
                onRefresh={onRefresh ?? (() => {})}
              />
            )}
            <AnalyticFields analyticId={analyticId} />
          </Box>
        )}
      </Box>
    </Box>
  )
}
