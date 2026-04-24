import { useState, useEffect, useCallback, useMemo } from 'react'
import { useTranslation } from 'react-i18next'
import { SimpleTreeView } from '@mui/x-tree-view/SimpleTreeView'
import { TreeItem } from '@mui/x-tree-view/TreeItem'
import { IconButton, Tooltip, Dialog, DialogTitle, DialogContent, DialogContentText, DialogActions, Button } from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import FolderOutlined from '@mui/icons-material/FolderOutlined'
import DescriptionOutlined from '@mui/icons-material/DescriptionOutlined'
import CategoryOutlined from '@mui/icons-material/CategoryOutlined'
import LockOutlined from '@mui/icons-material/LockOutlined'
import LockOpenOutlined from '@mui/icons-material/LockOpenOutlined'
import * as Icons from '@mui/icons-material'
import * as api from '../api'
import { currentLang } from '../i18n'
import type { Model, Sheet, Analytic, TreeSelection } from '../types'

interface Props {
  selection: TreeSelection | null
  onSelect: (sel: TreeSelection | null) => void
  refreshKey: number
  expandAfterCreate?: { modelId: string; folder: string; selectId: string; selectType: string } | null
  onCreated?: (info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => void
  sheetsOnly?: boolean
  currentUserId?: string
  isAdmin?: boolean
  onRefresh?: () => void
}

interface ModelTree {
  model: Model
  sheets: (Sheet & { can_edit?: boolean })[]
  analytics: Analytic[]
}

export default function LeftPanel({ selection, onSelect, refreshKey, expandAfterCreate, onCreated, sheetsOnly, currentUserId, isAdmin, onRefresh }: Props) {
  const { t } = useTranslation()
  const [trees, setTrees] = useState<ModelTree[]>([])
  const [search, setSearch] = useState('')
  const [expanded, setExpanded] = useState<string[]>([])
  const [dragSheet, setDragSheet] = useState<{ modelId: string; sheetId: string } | null>(null)
  const [dragOverSheet, setDragOverSheet] = useState<string | null>(null)
  const [deleteConfirm, setDeleteConfirm] = useState<{ id: string; name: string } | null>(null)
  // Translation map: "entity_type:entity_id:field" → translated value
  const [trMap, setTrMap] = useState<Record<string, string>>({})

  const load = useCallback(async () => {
    if (sheetsOnly && currentUserId) {
      const accessible = await api.getAccessibleSheets(currentUserId)
      const treesData: ModelTree[] = accessible.map((m: any) => ({
        model: { id: m.id, name: m.name } as Model,
        sheets: m.sheets.map((s: any) => ({ id: s.id, name: s.name, excel_code: s.excel_code, can_edit: s.can_edit, model_id: m.id, created_at: '', updated_at: '' } as Sheet & { can_edit?: boolean })),
        analytics: [],
      }))
      setTrees(treesData)
      setExpanded(prev => {
        const modelIds = treesData.map(t => `model:${t.model.id}`)
        const next = [...new Set([...prev, ...modelIds])]
        return next
      })
      // Load translations for all models
      const lang = currentLang()
      for (const td of treesData) {
        try {
          const tr = await api.getModelTranslations(td.model.id, lang)
          setTrMap(prev => ({ ...prev, ...tr }))
        } catch { /* ignore */ }
      }
    } else {
      const models = await api.listModels()
      const treesData: ModelTree[] = await Promise.all(
        models.map(async m => {
          const tree = await api.getModelTree(m.id)
          return { model: m, sheets: tree.sheets || [], analytics: tree.analytics || [] }
        })
      )
      setTrees(treesData)
      // Load translations
      const lang = currentLang()
      for (const td of treesData) {
        try {
          const tr = await api.getModelTranslations(td.model.id, lang)
          setTrMap(prev => ({ ...prev, ...tr }))
        } catch { /* ignore */ }
      }
    }
  }, [sheetsOnly, currentUserId])

  useEffect(() => { load() }, [load, refreshKey])

  // Reload only translations on language change (don't rebuild the tree)
  useEffect(() => {
    const handler = async () => {
      const lang = currentLang()
      for (const td of trees) {
        try {
          const tr = await api.getModelTranslations(td.model.id, lang)
          setTrMap(prev => ({ ...prev, ...tr }))
        } catch { /* ignore */ }
      }
    }
    window.addEventListener('pebble:langChange', handler)
    return () => window.removeEventListener('pebble:langChange', handler)
  }, [trees])

  // Auto-expand after create
  useEffect(() => {
    if (expandAfterCreate) {
      const { modelId, folder } = expandAfterCreate
      setExpanded(prev => {
        const needed = [`model:${modelId}`, `${folder}-folder:${modelId}`]
        const next = [...prev]
        for (const id of needed) {
          if (!next.includes(id)) next.push(id)
        }
        return next
      })
    }
  }, [expandAfterCreate])

  /** Get translated name or fallback to original */
  const tr = (entityType: string, entityId: string, fallback: string) => {
    const key = `${entityType}:${entityId}:name`
    return trMap[key] || fallback
  }

  const handleAdd = async () => {
    const m = await api.createModel(t('left.newModel'))
    await load()
    setExpanded(prev => [...prev, `model:${m.id}`])
    onSelect({ type: 'model', id: m.id, modelId: m.id })
  }

  const handleDeleteModel = async (e: React.MouseEvent, id: string, name?: string) => {
    e.stopPropagation()
    setDeleteConfirm({ id, name: name || t('left.noName') })
  }

  const confirmDeleteModel = async () => {
    if (!deleteConfirm) return
    const { id } = deleteConfirm
    setDeleteConfirm(null)
    await api.deleteModel(id)
    if (selection?.modelId === id) onSelect(null)
    load()
  }

  const handleAddSheet = async (e: React.MouseEvent, modelId: string) => {
    e.stopPropagation()
    const s = await api.createSheet({ model_id: modelId, name: t('left.newSheet') })
    await load()
    onCreated?.({ modelId, folder: 'sheets', id: s.id, type: 'sheet' })
  }

  const handleAddAnalytic = async (e: React.MouseEvent, modelId: string) => {
    e.stopPropagation()
    const a = await api.createAnalytic({ model_id: modelId, name: t('left.newAnalytic') })
    await load()
    onCreated?.({ modelId, folder: 'analytics', id: a.id, type: 'analytic' })
  }

  const handleDeleteSheet = async (e: React.MouseEvent, id: string) => {
    e.stopPropagation()
    await api.deleteSheet(id)
    if (selection?.id === id) onSelect(null)
    load()
  }

  const handleDeleteAnalytic = async (e: React.MouseEvent, id: string) => {
    e.stopPropagation()
    await api.deleteAnalytic(id)
    if (selection?.id === id) onSelect(null)
    load()
  }

  const handleItemSelect = (_e: React.SyntheticEvent, itemId: string | null) => {
    if (!itemId) return
    const parts = itemId.split(':')
    if (parts[0] === 'model') onSelect({ type: 'model', id: parts[1], modelId: parts[1] })
    else if (parts[0] === 'sheet') onSelect({ type: 'sheet', id: parts[1], modelId: parts[2] })
    else if (parts[0] === 'analytic') onSelect({ type: 'analytic', id: parts[1], modelId: parts[2] })
  }

  const handleExpandChange = (_e: React.SyntheticEvent, ids: string[]) => {
    setExpanded(ids)
  }

  const handleSheetDrop = async (modelId: string, targetSheetId: string) => {
    if (!dragSheet || dragSheet.modelId !== modelId || dragSheet.sheetId === targetSheetId) return
    const tree = trees.find(t => t.model.id === modelId)
    if (!tree) return
    const ids = tree.sheets.map(s => s.id)
    const fromIdx = ids.indexOf(dragSheet.sheetId)
    const toIdx = ids.indexOf(targetSheetId)
    if (fromIdx === -1 || toIdx === -1) return
    ids.splice(fromIdx, 1)
    ids.splice(toIdx, 0, dragSheet.sheetId)
    await api.reorderSheets(modelId, ids)
    load()
  }

  const getIcon = (iconName: string) => {
    if (!iconName) return <CategoryOutlined sx={{ fontSize: 18, opacity: 0.6 }} />
    const Icon = (Icons as any)[iconName]
    return Icon ? <Icon sx={{ fontSize: 18, opacity: 0.6 }} /> : <CategoryOutlined sx={{ fontSize: 18, opacity: 0.6 }} />
  }

  const q = search.toLowerCase()
  const selectedItemId = selection ? `${selection.type}:${selection.id}:${selection.modelId}` : ''

  // Sheets-only mode: flat list of models > sheets (no folders, no analytics)
  if (sheetsOnly) {
    return (
      <div className="panel-left">
        <div className="panel-left-toolbar">
          <input placeholder={t('left.search')} value={search} onChange={e => setSearch(e.target.value)} />
        </div>
        <div className="panel-left-tree">
          <SimpleTreeView
            selectedItems={selectedItemId}
            onSelectedItemsChange={handleItemSelect}
            expandedItems={expanded}
            onExpandedItemsChange={handleExpandChange}
          >
            {trees.map(({ model, sheets }) => {
              const modelName = tr('model', model.id, model.name) || t('left.noName')
              const mMatch = !q || modelName.toLowerCase().includes(q)
              const filteredSheets = sheets.filter(s => {
                const sName = tr('sheet', s.id, s.name)
                return !q || mMatch || sName.toLowerCase().includes(q)
              })
              if (!mMatch && filteredSheets.length === 0) return null

              return (
                <TreeItem
                  key={model.id}
                  itemId={`model:${model.id}`}
                  label={
                    <div className="tree-item-label">
                      <span style={{ fontWeight: 600 }}>{modelName}</span>
                      {isAdmin && (
                        <span className="actions">
                          <Tooltip title={t('left.deleteModel')}>
                            <IconButton size="small" onClick={e => handleDeleteModel(e, model.id, modelName)}>
                              <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                            </IconButton>
                          </Tooltip>
                        </span>
                      )}
                    </div>
                  }
                >
                  {filteredSheets.map(s => {
                    const sheetName = tr('sheet', s.id, s.name) || t('left.noName')
                    const isLocked = !!(s as any).locked
                    return (
                      <TreeItem
                        key={s.id}
                        itemId={`sheet:${s.id}:${model.id}`}
                        label={
                          <div className="tree-item-label">
                            <DescriptionOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                            {(s as any).excel_code && (
                              <span style={{ fontSize: 10, background: '#e3f2fd', color: '#1565c0', padding: '1px 4px', borderRadius: 3, fontWeight: 600, flexShrink: 0 }}>
                                {(s as any).excel_code}
                              </span>
                            )}
                            <span style={{ color: (s.can_edit === false || isLocked) ? '#999' : undefined }}>{sheetName}</span>
                            {s.can_edit === false && !isLocked && <LockOutlined sx={{ fontSize: 12, color: '#ccc', ml: 'auto' }} />}
                            {isLocked && (
                              <Tooltip title={isAdmin ? t('left.unlockSheet', 'Разблокировать лист') : t('left.sheetLocked', 'Лист заблокирован')}>
                                <span style={{ marginLeft: 'auto', display: 'flex' }}>
                                  <LockOutlined
                                    sx={{ fontSize: 14, color: '#ef5350', cursor: isAdmin ? 'pointer' : 'default' }}
                                    onClick={isAdmin ? (e) => {
                                      e.stopPropagation()
                                      api.toggleSheetLock(s.id).then(() => {
                                        setTrees(prev => prev.map(t => ({
                                          ...t,
                                          sheets: t.sheets.map(sh => sh.id === s.id ? { ...sh, locked: 0 } : sh)
                                        })))
                                        onRefresh?.()
                                      })
                                    } : undefined}
                                  />
                                </span>
                              </Tooltip>
                            )}
                            {!isLocked && isAdmin && (
                              <Tooltip title={t('left.lockSheet', 'Заблокировать лист')}>
                                <span style={{ marginLeft: isLocked || s.can_edit === false ? 0 : 'auto', display: 'flex' }}>
                                  <LockOpenOutlined
                                    sx={{ fontSize: 14, color: '#ccc', cursor: 'pointer', '&:hover': { color: '#999' } }}
                                    onClick={(e) => {
                                      e.stopPropagation()
                                      api.toggleSheetLock(s.id).then(() => {
                                        setTrees(prev => prev.map(t => ({
                                          ...t,
                                          sheets: t.sheets.map(sh => sh.id === s.id ? { ...sh, locked: 1 } : sh)
                                        })))
                                        onRefresh?.()
                                      })
                                    }}
                                  />
                                </span>
                              </Tooltip>
                            )}
                          </div>
                        }
                      />
                    )
                  })}
                </TreeItem>
              )
            })}
          </SimpleTreeView>
        </div>
        <Dialog open={!!deleteConfirm} onClose={() => setDeleteConfirm(null)}>
          <DialogTitle>{t('left.deleteModelTitle')}</DialogTitle>
          <DialogContent>
            <DialogContentText>
              {t('left.deleteModelText', { name: deleteConfirm?.name })}
            </DialogContentText>
          </DialogContent>
          <DialogActions>
            <Button onClick={() => setDeleteConfirm(null)}>{t('common.cancel')}</Button>
            <Button onClick={confirmDeleteModel} color="error" variant="contained">{t('common.delete')}</Button>
          </DialogActions>
        </Dialog>
      </div>
    )
  }

  return (
    <div className="panel-left">
      <div className="panel-left-toolbar">
        <input placeholder={t('left.search')} value={search} onChange={e => setSearch(e.target.value)} />
        <Tooltip title={t('left.addModel')}>
          <IconButton size="small" onClick={handleAdd}><AddOutlined fontSize="small" /></IconButton>
        </Tooltip>
      </div>
      <div className="panel-left-tree">
        <SimpleTreeView
          selectedItems={selectedItemId}
          onSelectedItemsChange={handleItemSelect}
          expandedItems={expanded}
          onExpandedItemsChange={handleExpandChange}
        >
          {trees.map(({ model, sheets, analytics }) => {
            const modelName = tr('model', model.id, model.name) || t('left.noName')
            const mMatch = !q || modelName.toLowerCase().includes(q)
            const filteredSheets = sheets.filter(s => {
              const sName = tr('sheet', s.id, s.name)
              return !q || mMatch || sName.toLowerCase().includes(q)
            })
            const filteredAnalytics = analytics.filter(a => {
              const aName = tr('analytic', a.id, a.name)
              return !q || mMatch || aName.toLowerCase().includes(q)
            })
            if (!mMatch && filteredSheets.length === 0 && filteredAnalytics.length === 0) return null

            return (
              <TreeItem
                key={model.id}
                itemId={`model:${model.id}`}
                label={
                  <div className="tree-item-label">
                    <span>{modelName}</span>
                    <span className="actions">
                      <Tooltip title={t('left.deleteModel')}>
                        <IconButton size="small" onClick={e => handleDeleteModel(e, model.id, modelName)}>
                          <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                        </IconButton>
                      </Tooltip>
                    </span>
                  </div>
                }
              >
                <TreeItem
                  itemId={`sheets-folder:${model.id}`}
                  label={
                    <div className="tree-item-label">
                      <FolderOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                      <span>{t('left.sheets')}</span>
                      <span className="actions">
                        <IconButton size="small" onClick={e => handleAddSheet(e, model.id)}>
                          <AddOutlined sx={{ fontSize: 16 }} />
                        </IconButton>
                      </span>
                    </div>
                  }
                >
                  {filteredSheets.map(s => {
                    const sheetName = tr('sheet', s.id, s.name) || t('left.noName')
                    return (
                      <TreeItem
                        key={s.id}
                        itemId={`sheet:${s.id}:${model.id}`}
                        label={
                          <div
                            className="tree-item-label"
                            draggable
                            onDragStart={e => { e.stopPropagation(); setDragSheet({ modelId: model.id, sheetId: s.id }) }}
                            onDragOver={e => { e.preventDefault(); e.stopPropagation(); setDragOverSheet(s.id) }}
                            onDragLeave={() => setDragOverSheet(null)}
                            onDrop={e => { e.preventDefault(); e.stopPropagation(); handleSheetDrop(model.id, s.id); setDragSheet(null); setDragOverSheet(null) }}
                            onDragEnd={() => { setDragSheet(null); setDragOverSheet(null) }}
                            style={{
                              opacity: dragSheet?.sheetId === s.id ? 0.4 : 1,
                              borderTop: dragOverSheet === s.id && dragSheet?.sheetId !== s.id ? '2px solid #1976d2' : '2px solid transparent',
                            }}
                          >
                            <DescriptionOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                            {(s as any).excel_code && (
                              <span style={{ fontSize: 10, background: '#e3f2fd', color: '#1565c0', padding: '1px 4px', borderRadius: 3, fontWeight: 600, flexShrink: 0 }}>
                                {(s as any).excel_code}
                              </span>
                            )}
                            <span style={{ color: (s as any).locked ? '#999' : undefined }}>{sheetName}</span>
                            {!!(s as any).locked && <LockOutlined sx={{ fontSize: 13, color: '#ef5350' }} />}
                            <span className="actions">
                              <IconButton size="small" onClick={e => handleDeleteSheet(e, s.id)}>
                                <DeleteOutlineOutlined sx={{ fontSize: 14 }} />
                              </IconButton>
                            </span>
                          </div>
                        }
                      />
                    )
                  })}
                </TreeItem>

                <TreeItem
                  itemId={`analytics-folder:${model.id}`}
                  label={
                    <div className="tree-item-label">
                      <FolderOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                      <span>{t('left.analytics')}</span>
                      <span className="actions">
                        <IconButton size="small" onClick={e => handleAddAnalytic(e, model.id)}>
                          <AddOutlined sx={{ fontSize: 16 }} />
                        </IconButton>
                      </span>
                    </div>
                  }
                >
                  {filteredAnalytics.map(a => {
                    const aName = tr('analytic', a.id, a.name) || t('left.noName')
                    return (
                      <TreeItem
                        key={a.id}
                        itemId={`analytic:${a.id}:${model.id}`}
                        label={
                          <div className="tree-item-label">
                            {getIcon(a.icon)}
                            <span>{aName}</span>
                            <span className="actions">
                              <IconButton size="small" onClick={e => handleDeleteAnalytic(e, a.id)}>
                                <DeleteOutlineOutlined sx={{ fontSize: 14 }} />
                              </IconButton>
                            </span>
                          </div>
                        }
                      />
                    )
                  })}
                </TreeItem>
              </TreeItem>
            )
          })}
        </SimpleTreeView>
      </div>
      <Dialog open={!!deleteConfirm} onClose={() => setDeleteConfirm(null)}>
        <DialogTitle>{t('left.deleteModelTitle')}</DialogTitle>
        <DialogContent>
          <DialogContentText>
            {t('left.deleteModelText', { name: deleteConfirm?.name })}
          </DialogContentText>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setDeleteConfirm(null)}>{t('common.cancel')}</Button>
          <Button onClick={confirmDeleteModel} color="error" variant="contained">{t('common.delete')}</Button>
        </DialogActions>
      </Dialog>
    </div>
  )
}
