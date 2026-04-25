import { useState, useEffect, useRef } from 'react'
import {
  Box, Typography, TextField, List, ListItem, ListItemIcon, ListItemText,
  IconButton, Menu, MenuItem, Tooltip, Radio,
} from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import * as Icons from '@mui/icons-material'
import { useTranslation } from 'react-i18next'
import * as api from '../../api'
import type { Sheet, SheetAnalytic, Analytic } from '../../types'
import { usePending } from '../../store/PendingContext'
import RecordOrganizer from './RecordOrganizer'

interface Props {
  sheetId: string
  modelId: string
}

export default function SheetSettings({ sheetId, modelId }: Props) {
  const { t } = useTranslation()
  const [sheet, setSheet] = useState<Sheet | null>(null)
  const [bindings, setBindings] = useState<SheetAnalytic[]>([])
  const [allAnalytics, setAllAnalytics] = useState<Analytic[]>([])
  const [mainAnalyticId, setMainAnalyticId] = useState<string | null>(null)
  const [menuAnchor, setMenuAnchor] = useState<HTMLElement | null>(null)
  const { addOp, getOverrides } = usePending()
  const dragIdx = useRef<number | null>(null)
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null)

  const load = async () => {
    const [sheets, sa, analytics, main] = await Promise.all([
      api.listSheets(modelId),
      api.listSheetAnalytics(sheetId),
      api.listAnalytics(modelId),
      api.getMainAnalytic(sheetId),
    ])
    const found = sheets.find(s => s.id === sheetId) || null
    if (found) {
      const overrides = getOverrides(`sheet:${sheetId}`)
      if (overrides) Object.assign(found, overrides)
    }
    setSheet(found)
    setBindings(sa)
    setAllAnalytics(analytics)
    setMainAnalyticId(main.analytic_id)
  }

  useEffect(() => { load() }, [sheetId, modelId])

  if (!sheet) return null

  const changeName = (name: string) => {
    setSheet({ ...sheet, name })
    addOp({
      key: `sheet:${sheetId}`,
      type: 'updateSheet',
      id: sheetId,
      data: { name },
    })
  }

  const boundIds = new Set(bindings.map(b => b.analytic_id))
  const available = allAnalytics.filter(a => !boundIds.has(a.id))

  const handleAddAnalytic = async (analyticId: string) => {
    setMenuAnchor(null)
    await api.addSheetAnalytic(sheetId, { analytic_id: analyticId, sort_order: bindings.length })
    load()
  }

  const handleRemove = async (saId: string) => {
    await api.removeSheetAnalytic(sheetId, saId)
    load()
  }

  const handleSetMain = async (analyticId: string) => {
    setMainAnalyticId(analyticId)  // optimistic
    try {
      await api.setMainAnalytic(sheetId, analyticId)
    } catch {
      load()  // revert on error
    }
  }

  const analyticById = (aid: string) => allAnalytics.find(a => a.id === aid)

  const handleDragStart = (idx: number) => {
    dragIdx.current = idx
  }

  const handleDragOver = (e: React.DragEvent, idx: number) => {
    e.preventDefault()
    setDragOverIdx(idx)
  }

  const handleDrop = async (idx: number) => {
    const from = dragIdx.current
    if (from === null || from === idx) {
      dragIdx.current = null
      setDragOverIdx(null)
      return
    }
    const ordered = [...bindings]
    const [moved] = ordered.splice(from, 1)
    ordered.splice(idx, 0, moved)
    setBindings(ordered)
    dragIdx.current = null
    setDragOverIdx(null)
    await api.reorderSheetAnalytics(sheetId, ordered.map(b => b.id))
    load()
  }

  const handleDragEnd = () => {
    dragIdx.current = null
    setDragOverIdx(null)
  }

  const getIcon = (iconName?: string) => {
    if (!iconName) return <Icons.CategoryOutlined fontSize="small" />
    const Icon = (Icons as any)[iconName]
    return Icon ? <Icon fontSize="small" /> : <Icons.CategoryOutlined fontSize="small" />
  }

  return (
    <Box sx={{ maxWidth: 500 }}>
      <Typography variant="h6" sx={{ mb: 2 }}>{t('sheet.title')}</Typography>
      <TextField
        label={t('fields.name')} fullWidth value={sheet.name}
        onChange={e => changeName(e.target.value)} sx={{ mb: 3 }}
      />

      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
        <Typography variant="subtitle1">{t('sheet.linkedAnalytics')}</Typography>
        <Tooltip title={t('sheet.addAnalytic')}>
          <span>
            <IconButton
              size="small"
              disabled={available.length === 0}
              onClick={e => setMenuAnchor(e.currentTarget)}
            >
              <AddOutlined fontSize="small" />
            </IconButton>
          </span>
        </Tooltip>
        <Menu anchorEl={menuAnchor} open={!!menuAnchor} onClose={() => setMenuAnchor(null)}>
          {available.map(a => (
            <MenuItem key={a.id} onClick={() => handleAddAnalytic(a.id)}>
              {getIcon(a.icon)}
              <Typography sx={{ ml: 1 }}>{a.name}</Typography>
            </MenuItem>
          ))}
        </Menu>
      </Box>

      <List dense>
        {bindings.map((b, i) => (
          <Box key={b.id}>
            <ListItem
              draggable
              onDragStart={() => handleDragStart(i)}
              onDragOver={e => handleDragOver(e, i)}
              onDrop={() => handleDrop(i)}
              onDragEnd={handleDragEnd}
              sx={{
                cursor: 'grab',
                borderTop: dragOverIdx === i ? '2px solid #1976d2' : '2px solid transparent',
                opacity: dragIdx.current === i ? 0.5 : 1,
                '&:hover': { bgcolor: '#f5f5f5' },
              }}
              secondaryAction={
                <IconButton size="small" onClick={() => handleRemove(b.id)}>
                  <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                </IconButton>
              }
            >
              <ListItemIcon sx={{ minWidth: 28 }}>
                <DragIndicatorOutlined sx={{ fontSize: 16, color: '#bbb', cursor: 'grab' }} />
              </ListItemIcon>
              <ListItemIcon sx={{ minWidth: 32 }}>
                {getIcon(b.analytic_icon)}
              </ListItemIcon>
              <ListItemText primary={b.analytic_name || t('sheet.analyticLabel')} />
              {(() => {
                const a = analyticById(b.analytic_id)
                if (a && a.is_periods) return null
                const isMain = b.analytic_id === mainAnalyticId
                return (
                  <Tooltip title={isMain ? t('sheet.mainAnalytic') : t('sheet.makeMain')}>
                    <Radio
                      size="small"
                      checked={isMain}
                      onChange={() => handleSetMain(b.analytic_id)}
                      sx={{ p: 0.5, mr: 4 }}
                    />
                  </Tooltip>
                )
              })()}
            </ListItem>
            {/* Record organizer for all analytics */}
            {(() => {
              const a = analyticById(b.analytic_id)
              if (!a) return null
              return (
                <RecordOrganizer
                  sheetId={sheetId}
                  saId={b.id}
                  analyticId={b.analytic_id}
                  isPeriods={!!a.is_periods}
                  initialVisible={b.visible_record_ids || null}
                  onSaved={load}
                />
              )
            })()}
          </Box>
        ))}
      </List>

      {bindings.length === 0 && (
        <Typography variant="body2" color="textSecondary">
          {t('sheet.noAnalytics')}
        </Typography>
      )}

    </Box>
  )
}
