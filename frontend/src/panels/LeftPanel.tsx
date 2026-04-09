import { useState, useEffect, useCallback, useMemo } from 'react'
import { SimpleTreeView } from '@mui/x-tree-view/SimpleTreeView'
import { TreeItem } from '@mui/x-tree-view/TreeItem'
import { IconButton, Tooltip } from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import FolderOutlined from '@mui/icons-material/FolderOutlined'
import DescriptionOutlined from '@mui/icons-material/DescriptionOutlined'
import CategoryOutlined from '@mui/icons-material/CategoryOutlined'
import * as Icons from '@mui/icons-material'
import * as api from '../api'
import type { Model, Sheet, Analytic, TreeSelection } from '../types'

interface Props {
  selection: TreeSelection | null
  onSelect: (sel: TreeSelection | null) => void
  refreshKey: number
  expandAfterCreate?: { modelId: string; folder: string; selectId: string; selectType: string } | null
  onCreated?: (info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => void
}

interface ModelTree {
  model: Model
  sheets: Sheet[]
  analytics: Analytic[]
}

export default function LeftPanel({ selection, onSelect, refreshKey, expandAfterCreate, onCreated }: Props) {
  const [trees, setTrees] = useState<ModelTree[]>([])
  const [search, setSearch] = useState('')
  const [expanded, setExpanded] = useState<string[]>([])

  const load = useCallback(async () => {
    const models = await api.listModels()
    const treesData: ModelTree[] = await Promise.all(
      models.map(async m => {
        const tree = await api.getModelTree(m.id)
        return { model: m, sheets: tree.sheets || [], analytics: tree.analytics || [] }
      })
    )
    setTrees(treesData)
  }, [])

  useEffect(() => { load() }, [load, refreshKey])

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

  const handleAdd = async () => {
    const m = await api.createModel('Новая модель')
    await load()
    setExpanded(prev => [...prev, `model:${m.id}`])
    onSelect({ type: 'model', id: m.id, modelId: m.id })
  }

  const handleDeleteModel = async (e: React.MouseEvent, id: string) => {
    e.stopPropagation()
    await api.deleteModel(id)
    if (selection?.modelId === id) onSelect(null)
    load()
  }

  const handleAddSheet = async (e: React.MouseEvent, modelId: string) => {
    e.stopPropagation()
    const s = await api.createSheet({ model_id: modelId, name: 'Новый лист' })
    await load()
    onCreated?.({ modelId, folder: 'sheets', id: s.id, type: 'sheet' })
  }

  const handleAddAnalytic = async (e: React.MouseEvent, modelId: string) => {
    e.stopPropagation()
    const a = await api.createAnalytic({ model_id: modelId, name: 'Новая аналитика' })
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

  const getIcon = (iconName: string) => {
    if (!iconName) return <CategoryOutlined sx={{ fontSize: 18, opacity: 0.6 }} />
    const Icon = (Icons as any)[iconName]
    return Icon ? <Icon sx={{ fontSize: 18, opacity: 0.6 }} /> : <CategoryOutlined sx={{ fontSize: 18, opacity: 0.6 }} />
  }

  const q = search.toLowerCase()
  const selectedItemId = selection ? `${selection.type}:${selection.id}:${selection.modelId}` : ''

  return (
    <div className="panel-left">
      <div className="panel-left-toolbar">
        <input placeholder="Поиск..." value={search} onChange={e => setSearch(e.target.value)} />
        <Tooltip title="Добавить модель">
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
            const mMatch = !q || model.name.toLowerCase().includes(q)
            const filteredSheets = sheets.filter(s => !q || mMatch || s.name.toLowerCase().includes(q))
            const filteredAnalytics = analytics.filter(a => !q || mMatch || a.name.toLowerCase().includes(q))
            if (!mMatch && filteredSheets.length === 0 && filteredAnalytics.length === 0) return null

            return (
              <TreeItem
                key={model.id}
                itemId={`model:${model.id}`}
                label={
                  <div className="tree-item-label">
                    <span>{model.name || 'Без названия'}</span>
                    <span className="actions">
                      <IconButton size="small" onClick={e => handleDeleteModel(e, model.id)}>
                        <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                      </IconButton>
                    </span>
                  </div>
                }
              >
                <TreeItem
                  itemId={`sheets-folder:${model.id}`}
                  label={
                    <div className="tree-item-label">
                      <FolderOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                      <span>Листы</span>
                      <span className="actions">
                        <IconButton size="small" onClick={e => handleAddSheet(e, model.id)}>
                          <AddOutlined sx={{ fontSize: 16 }} />
                        </IconButton>
                      </span>
                    </div>
                  }
                >
                  {filteredSheets.map(s => (
                    <TreeItem
                      key={s.id}
                      itemId={`sheet:${s.id}:${model.id}`}
                      label={
                        <div className="tree-item-label">
                          <DescriptionOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                          <span>{s.name || 'Без названия'}</span>
                          <span className="actions">
                            <IconButton size="small" onClick={e => handleDeleteSheet(e, s.id)}>
                              <DeleteOutlineOutlined sx={{ fontSize: 14 }} />
                            </IconButton>
                          </span>
                        </div>
                      }
                    />
                  ))}
                </TreeItem>

                <TreeItem
                  itemId={`analytics-folder:${model.id}`}
                  label={
                    <div className="tree-item-label">
                      <FolderOutlined sx={{ fontSize: 16, opacity: 0.5 }} />
                      <span>Аналитики</span>
                      <span className="actions">
                        <IconButton size="small" onClick={e => handleAddAnalytic(e, model.id)}>
                          <AddOutlined sx={{ fontSize: 16 }} />
                        </IconButton>
                      </span>
                    </div>
                  }
                >
                  {filteredAnalytics.map(a => (
                    <TreeItem
                      key={a.id}
                      itemId={`analytic:${a.id}:${model.id}`}
                      label={
                        <div className="tree-item-label">
                          {getIcon(a.icon)}
                          <span>{a.name || 'Без названия'}</span>
                          <span className="actions">
                            <IconButton size="small" onClick={e => handleDeleteAnalytic(e, a.id)}>
                              <DeleteOutlineOutlined sx={{ fontSize: 14 }} />
                            </IconButton>
                          </span>
                        </div>
                      }
                    />
                  ))}
                </TreeItem>
              </TreeItem>
            )
          })}
        </SimpleTreeView>
      </div>
    </div>
  )
}
