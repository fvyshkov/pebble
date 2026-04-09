import { useState, useEffect, useRef } from 'react'
import {
  Box, Typography, TextField, List, ListItem, ListItemIcon, ListItemText,
  IconButton, Menu, MenuItem, Tooltip, Button,
  Table, TableHead, TableBody, TableRow, TableCell, Checkbox,
} from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import * as Icons from '@mui/icons-material'
import * as api from '../../api'
import type { Sheet, SheetAnalytic, Analytic } from '../../types'
import { usePending } from '../../store/PendingContext'

interface Props {
  sheetId: string
  modelId: string
}

export default function SheetSettings({ sheetId, modelId }: Props) {
  const [sheet, setSheet] = useState<Sheet | null>(null)
  const [bindings, setBindings] = useState<SheetAnalytic[]>([])
  const [allAnalytics, setAllAnalytics] = useState<Analytic[]>([])
  const [menuAnchor, setMenuAnchor] = useState<HTMLElement | null>(null)
  const { addOp, getOverrides } = usePending()
  const dragIdx = useRef<number | null>(null)
  const [dragOverIdx, setDragOverIdx] = useState<number | null>(null)

  const load = async () => {
    const [sheets, sa, analytics] = await Promise.all([
      api.listSheets(modelId),
      api.listSheetAnalytics(sheetId),
      api.listAnalytics(modelId),
    ])
    const found = sheets.find(s => s.id === sheetId) || null
    if (found) {
      const overrides = getOverrides(`sheet:${sheetId}`)
      if (overrides) Object.assign(found, overrides)
    }
    setSheet(found)
    setBindings(sa)
    setAllAnalytics(analytics)
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
      <Typography variant="h6" sx={{ mb: 2 }}>Лист</Typography>
      <TextField
        label="Название" fullWidth value={sheet.name}
        onChange={e => changeName(e.target.value)} sx={{ mb: 3 }}
      />

      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1 }}>
        <Typography variant="subtitle1">Привязанные аналитики</Typography>
        <Tooltip title="Добавить аналитику">
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
          <ListItem
            key={b.id}
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
            <ListItemText primary={b.analytic_name || 'Аналитика'} />
          </ListItem>
        ))}
      </List>

      {bindings.length === 0 && (
        <Typography variant="body2" color="textSecondary">
          Нет привязанных аналитик. Добавьте аналитику к листу.
        </Typography>
      )}

      <SheetPermissions sheetId={sheetId} />
    </Box>
  )
}

// ─── Inline permissions grid ───
interface PermRow { user_id: string; username: string; can_view: number; can_edit: number }

function SheetPermissions({ sheetId }: { sheetId: string }) {
  const [perms, setPerms] = useState<PermRow[]>([])

  useEffect(() => {
    api.getSheetPermissions(sheetId).then(setPerms)
  }, [sheetId])

  const handleToggle = async (userId: string, field: 'can_view' | 'can_edit', current: number) => {
    const row = perms.find(p => p.user_id === userId)
    if (!row) return
    await api.setSheetPermission(sheetId, {
      user_id: userId,
      can_view: field === 'can_view' ? !current : !!row.can_view,
      can_edit: field === 'can_edit' ? !current : !!row.can_edit,
    })
    api.getSheetPermissions(sheetId).then(setPerms)
  }

  if (perms.length === 0) return null

  return (
    <Box sx={{ mt: 4 }}>
      <Typography variant="subtitle1" sx={{ mb: 1 }}>Доступ к листу</Typography>
      <Table size="small" sx={{ maxWidth: 400, '& td, & th': { py: 0.25, px: 0.5 } }}>
        <TableHead>
          <TableRow>
            <TableCell>Пользователь</TableCell>
            <TableCell align="center" sx={{ fontSize: 11 }}>Просмотр</TableCell>
            <TableCell align="center" sx={{ fontSize: 11 }}>Редакт.</TableCell>
          </TableRow>
        </TableHead>
        <TableBody>
          {perms.map(p => (
            <TableRow key={p.user_id}>
              <TableCell sx={{ fontSize: 12 }}>{p.username}</TableCell>
              <TableCell align="center">
                <Checkbox size="small" checked={!!p.can_view}
                  onChange={() => handleToggle(p.user_id, 'can_view', p.can_view)} sx={{ p: 0 }} />
              </TableCell>
              <TableCell align="center">
                <Checkbox size="small" checked={!!p.can_edit}
                  onChange={() => handleToggle(p.user_id, 'can_edit', p.can_edit)} sx={{ p: 0 }} />
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </Box>
  )
}
