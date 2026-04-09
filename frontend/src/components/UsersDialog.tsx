import { useState, useEffect } from 'react'
import {
  Box, Typography, IconButton, TextField, Table, TableHead, TableBody,
  TableRow, TableCell, Button, Tooltip, Checkbox,
} from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import RestartAltOutlined from '@mui/icons-material/RestartAltOutlined'
import AddOutlined from '@mui/icons-material/AddOutlined'
import * as api from '../api'

interface Props {
  open: boolean
  onClose: () => void
}

export default function UsersDialog({ open, onClose }: Props) {
  const [users, setUsers] = useState<any[]>([])
  const [newName, setNewName] = useState('')

  const load = () => { api.listUsers().then(setUsers) }
  useEffect(() => { if (open) load() }, [open])

  if (!open) return null

  const handleAdd = async () => {
    const name = newName.trim()
    if (!name) return
    await api.createUser(name)
    setNewName('')
    load()
  }

  const handleDelete = async (id: string) => {
    await api.deleteUser(id)
    load()
  }

  const handleReset = async (id: string) => {
    const pw = prompt('Введите новый пароль:')
    if (!pw) return
    await api.resetPassword(id, pw)
    alert('Пароль изменён')
  }

  const handleToggleAdmin = async (id: string, current: boolean) => {
    await api.setAdmin(id, !current)
    load()
  }

  return (
    <Box sx={{
      position: 'fixed', inset: 0, zIndex: 1300, bgcolor: '#fff',
      display: 'flex', flexDirection: 'column',
    }}>
      <Box sx={{ display: 'flex', alignItems: 'center', px: 2, py: 1, borderBottom: '1px solid #e0e0e0', gap: 1 }}>
        <Typography variant="h6" sx={{ flex: 1 }}>Пользователи</Typography>
        <IconButton size="small" onClick={onClose}><CloseOutlined fontSize="small" /></IconButton>
      </Box>
      <Box sx={{ flex: 1, overflow: 'auto', p: 2, maxWidth: 700 }}>
        <Box sx={{ display: 'flex', gap: 1, mb: 2 }}>
          <TextField
            size="small" placeholder="Имя пользователя" value={newName}
            onChange={e => setNewName(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') handleAdd() }}
            sx={{ flex: 1 }}
          />
          <Button variant="outlined" startIcon={<AddOutlined />} onClick={handleAdd} disabled={!newName.trim()}>
            Добавить
          </Button>
        </Box>

        <Table size="small">
          <TableHead>
            <TableRow>
              <TableCell>Имя</TableCell>
              <TableCell align="center">Админ</TableCell>
              <TableCell>Создан</TableCell>
              <TableCell width={100}></TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {users.map(u => (
              <TableRow key={u.id} hover>
                <TableCell>{u.username}</TableCell>
                <TableCell align="center">
                  <Checkbox
                    size="small" checked={!!u.can_admin}
                    onChange={() => handleToggleAdmin(u.id, !!u.can_admin)}
                  />
                </TableCell>
                <TableCell>{u.created_at?.slice(0, 10)}</TableCell>
                <TableCell>
                  <Box sx={{ display: 'flex', gap: 0.5 }}>
                    <Tooltip title="Сбросить пароль">
                      <IconButton size="small" onClick={() => handleReset(u.id)}>
                        <RestartAltOutlined sx={{ fontSize: 16 }} />
                      </IconButton>
                    </Tooltip>
                    <Tooltip title="Удалить">
                      <IconButton size="small" onClick={() => handleDelete(u.id)}>
                        <DeleteOutlineOutlined sx={{ fontSize: 16 }} />
                      </IconButton>
                    </Tooltip>
                  </Box>
                </TableCell>
              </TableRow>
            ))}
            {users.length === 0 && (
              <TableRow><TableCell colSpan={4}><Typography variant="body2" color="textSecondary">Нет пользователей</Typography></TableCell></TableRow>
            )}
          </TableBody>
        </Table>
      </Box>
    </Box>
  )
}
