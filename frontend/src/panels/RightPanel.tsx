import { useState, useEffect } from 'react'
import {
  Box, Typography, Table, TableHead, TableBody, TableRow, TableCell, Checkbox,
} from '@mui/material'
import type { TreeSelection } from '../types'
import * as api from '../api'

interface Props {
  selection: TreeSelection | null
}

interface PermRow {
  user_id: string
  username: string
  can_view: number
  can_edit: number
}

export default function RightPanel({ selection }: Props) {
  const [perms, setPerms] = useState<PermRow[]>([])

  const sheetId = selection?.type === 'sheet' ? selection.id : null

  useEffect(() => {
    if (!sheetId) { setPerms([]); return }
    api.getSheetPermissions(sheetId).then(setPerms)
  }, [sheetId])

  if (!sheetId) return null

  const handleToggle = async (userId: string, field: 'can_view' | 'can_edit', current: number) => {
    const row = perms.find(p => p.user_id === userId)
    if (!row) return
    const data = {
      user_id: userId,
      can_view: field === 'can_view' ? !current : !!row.can_view,
      can_edit: field === 'can_edit' ? !current : !!row.can_edit,
    }
    await api.setSheetPermission(sheetId, data)
    api.getSheetPermissions(sheetId).then(setPerms)
  }

  if (perms.length === 0) return null

  return (
    <Box sx={{ p: 1, minWidth: 220 }}>
      <Typography variant="subtitle2" sx={{ mb: 1 }}>Доступ к листу</Typography>
      <Table size="small" sx={{ '& td, & th': { py: 0.25, px: 0.5 } }}>
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
