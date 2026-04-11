import { useState, useEffect, useCallback } from 'react'
import {
  Box, Typography, IconButton, TextField, Tooltip, Checkbox,
  FormControlLabel, Switch, Divider,
} from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import PersonAddOutlined from '@mui/icons-material/PersonAddOutlined'
import PersonOutlined from '@mui/icons-material/PersonOutlined'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import ChevronRightOutlined from '@mui/icons-material/ChevronRightOutlined'
import DescriptionOutlined from '@mui/icons-material/DescriptionOutlined'
import KeyOutlined from '@mui/icons-material/KeyOutlined'
import * as api from '../api'

interface Props {
  open: boolean
  onClose: () => void
}

interface PermModel {
  id: string; name: string
  sheets: { id: string; name: string; can_view: boolean; can_edit: boolean }[]
}

export default function UsersDialog({ open, onClose }: Props) {
  const [users, setUsers] = useState<any[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [perms, setPerms] = useState<PermModel[]>([])
  const [analyticPerms, setAnalyticPerms] = useState<any[]>([])
  const [username, setUsername] = useState('')
  const [isAdmin, setIsAdmin] = useState(false)
  const [createdAt, setCreatedAt] = useState('')
  const [expanded, setExpanded] = useState<Set<string>>(new Set())

  const loadUsers = useCallback(() => {
    api.listUsers().then(u => {
      setUsers(u)
      if (u.length > 0 && !selectedId) setSelectedId(u[0].id)
    })
  }, [])

  useEffect(() => { if (open) loadUsers() }, [open, loadUsers])

  const loadPerms = useCallback(() => {
    if (!selectedId) return
    api.getAllPermissions(selectedId).then(setPerms)
    api.getAnalyticPermissions(selectedId).then(setAnalyticPerms)
  }, [selectedId])

  useEffect(() => {
    if (!selectedId) return
    const u = users.find(x => x.id === selectedId)
    if (u) {
      setUsername(u.username)
      setIsAdmin(!!u.can_admin)
      setCreatedAt(u.created_at?.slice(0, 10) || '')
    }
    loadPerms()
  }, [selectedId, users, loadPerms])

  if (!open) return null

  const handleAdd = async () => {
    const u = await api.createUser('Новый пользователь')
    await loadUsers()
    setSelectedId(u.id)
  }

  const handleDelete = async () => {
    if (!selectedId) return
    if (!confirm('Удалить пользователя?')) return
    await api.deleteUser(selectedId)
    setSelectedId(null)
    loadUsers()
  }

  const handleSaveName = async () => {
    if (!selectedId || !username.trim()) return
    await api.updateUser(selectedId, username.trim())
    loadUsers()
  }

  const handleToggleAdmin = async (val: boolean) => {
    if (!selectedId) return
    setIsAdmin(val)
    await api.setAdmin(selectedId, val)
    loadUsers()
  }

  const handleResetPassword = async () => {
    if (!selectedId) return
    const pw = prompt('Новый пароль:')
    if (!pw) return
    await api.resetPassword(selectedId, pw)
    alert('Пароль изменён')
  }

  const handleSheetPerm = async (sheetId: string, field: 'can_view' | 'can_edit', value: boolean) => {
    if (!selectedId) return
    // Find current perms
    const model = perms.find(m => m.sheets.some(s => s.id === sheetId))
    const sheet = model?.sheets.find(s => s.id === sheetId)
    if (!sheet) return
    const canView = field === 'can_view' ? value : sheet.can_view
    const canEdit = field === 'can_edit' ? value : sheet.can_edit
    await api.setSheetPermission(sheetId, { user_id: selectedId, can_view: canView, can_edit: canEdit })
    loadPerms()
  }

  const handleModelPerm = async (modelId: string, field: 'can_view' | 'can_edit', value: boolean) => {
    if (!selectedId) return
    const model = perms.find(m => m.id === modelId)
    if (!model) return
    for (const s of model.sheets) {
      const canView = field === 'can_view' ? value : s.can_view
      const canEdit = field === 'can_edit' ? value : s.can_edit
      await api.setSheetPermission(s.id, { user_id: selectedId, can_view: canView, can_edit: canEdit })
    }
    loadPerms()
  }

  const modelChecked = (modelId: string, field: 'can_view' | 'can_edit') => {
    const model = perms.find(m => m.id === modelId)
    if (!model || model.sheets.length === 0) return false
    return model.sheets.every(s => s[field])
  }

  const modelIndeterminate = (modelId: string, field: 'can_view' | 'can_edit') => {
    const model = perms.find(m => m.id === modelId)
    if (!model || model.sheets.length === 0) return false
    const on = model.sheets.filter(s => s[field]).length
    return on > 0 && on < model.sheets.length
  }

  return (
    <Box sx={{ position: 'fixed', inset: 0, zIndex: 1300, bgcolor: '#fff', display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <Box sx={{ display: 'flex', alignItems: 'center', px: 2, py: 1, borderBottom: '1px solid #e0e0e0' }}>
        <Typography variant="h6" sx={{ flex: 1 }}>Пользователи</Typography>
        <IconButton size="small" onClick={onClose}><CloseOutlined fontSize="small" /></IconButton>
      </Box>

      <Box sx={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        {/* ─── Left: user list ─── */}
        <Box sx={{ width: 220, borderRight: '1px solid #e0e0e0', display: 'flex', flexDirection: 'column' }}>
          <Box sx={{ p: 1, borderBottom: '1px solid #f0f0f0', display: 'flex', gap: 0.5 }}>
            <Tooltip title="Добавить пользователя">
              <IconButton size="small" onClick={handleAdd}><PersonAddOutlined fontSize="small" /></IconButton>
            </Tooltip>
          </Box>
          <Box sx={{ flex: 1, overflow: 'auto' }}>
            {users.map(u => (
              <Box
                key={u.id}
                onClick={() => setSelectedId(u.id)}
                sx={{
                  display: 'flex', alignItems: 'center', gap: 1, px: 2, py: 1,
                  cursor: 'pointer', fontSize: 14,
                  bgcolor: u.id === selectedId ? '#e3f2fd' : 'transparent',
                  fontWeight: u.id === selectedId ? 600 : 400,
                  '&:hover': { bgcolor: u.id === selectedId ? '#e3f2fd' : '#f5f5f5' },
                }}
              >
                <PersonOutlined sx={{ fontSize: 18, opacity: 0.5 }} />
                {u.username}
                {!!u.can_admin && <Typography sx={{ fontSize: 10, color: '#1976d2', ml: 'auto' }}>админ</Typography>}
              </Box>
            ))}
          </Box>
        </Box>

        {/* ─── Center: details + permissions ─── */}
        {selectedId ? (
          <Box sx={{ flex: 1, overflow: 'auto', p: 3, maxWidth: 800 }}>
            {/* User attributes */}
            <Box sx={{ display: 'flex', gap: 2, alignItems: 'flex-start', mb: 3 }}>
              <TextField
                label="Имя" size="small" value={username}
                onChange={e => setUsername(e.target.value)}
                onBlur={handleSaveName}
                onKeyDown={e => { if (e.key === 'Enter') handleSaveName() }}
                sx={{ flex: 1, maxWidth: 300 }}
              />
              <FormControlLabel
                control={<Switch checked={isAdmin} onChange={e => handleToggleAdmin(e.target.checked)} size="small" />}
                label="Админ"
              />
              <Tooltip title="Сбросить пароль">
                <IconButton size="small" onClick={handleResetPassword}>
                  <KeyOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
              <Tooltip title="Удалить пользователя">
                <IconButton size="small" color="error" onClick={handleDelete}>
                  <DeleteOutlineOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
            </Box>
            <Typography variant="caption" color="textSecondary" sx={{ mb: 3, display: 'block' }}>
              Создан: {createdAt}
            </Typography>

            <Divider sx={{ mb: 2 }} />

            {/* Permissions tree */}
            <Typography variant="subtitle2" sx={{ mb: 1 }}>Доступ к листам</Typography>

            <Box sx={{ fontSize: 13 }}>
              {/* Header */}
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.5, borderBottom: '1px solid #e0e0e0', fontWeight: 600, color: '#666' }}>
                <Box sx={{ flex: 1 }}>Модель / Лист</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>Просмотр</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>Ввод</Box>
              </Box>

              {perms.map(model => {
                const isOpen = expanded.has(model.id)
                const toggle = () => setExpanded(prev => {
                  const next = new Set(prev)
                  if (next.has(model.id)) next.delete(model.id); else next.add(model.id)
                  return next
                })
                return (
                <Box key={model.id}>
                  {/* Model row */}
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.5, bgcolor: '#fafafa', '&:hover': { bgcolor: '#f0f0f0' }, cursor: 'pointer' }}
                    onClick={toggle}>
                    {isOpen ? <ExpandMoreOutlined sx={{ fontSize: 18, opacity: 0.5 }} /> : <ChevronRightOutlined sx={{ fontSize: 18, opacity: 0.5 }} />}
                    <Box sx={{ flex: 1, fontWeight: 600 }}>{model.name}</Box>
                    <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                      <Checkbox
                        size="small"
                        checked={modelChecked(model.id, 'can_view')}
                        indeterminate={modelIndeterminate(model.id, 'can_view')}
                        onChange={e => handleModelPerm(model.id, 'can_view', e.target.checked)}
                      />
                    </Box>
                    <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                      <Checkbox
                        size="small"
                        checked={modelChecked(model.id, 'can_edit')}
                        indeterminate={modelIndeterminate(model.id, 'can_edit')}
                        onChange={e => handleModelPerm(model.id, 'can_edit', e.target.checked)}
                      />
                    </Box>
                  </Box>

                  {/* Sheet rows (collapsible) */}
                  {isOpen && model.sheets.map(sheet => (
                    <Box key={sheet.id} sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.25, pl: 5, '&:hover': { bgcolor: '#f8f8f8' } }}>
                      <DescriptionOutlined sx={{ fontSize: 14, opacity: 0.4 }} />
                      <Box sx={{ flex: 1 }}>{sheet.name}</Box>
                      <Box sx={{ width: 80, textAlign: 'center' }}>
                        <Checkbox
                          size="small"
                          checked={sheet.can_view}
                          onChange={e => handleSheetPerm(sheet.id, 'can_view', e.target.checked)}
                        />
                      </Box>
                      <Box sx={{ width: 80, textAlign: 'center' }}>
                        <Checkbox
                          size="small"
                          checked={sheet.can_edit}
                          onChange={e => handleSheetPerm(sheet.id, 'can_edit', e.target.checked)}
                        />
                      </Box>
                    </Box>
                  ))}
                </Box>
                )
              })}

              {perms.length === 0 && (
                <Typography sx={{ py: 2, color: '#999', textAlign: 'center' }}>Нет моделей</Typography>
              )}
            </Box>

            <Divider sx={{ my: 2 }} />

            {/* Analytic record permissions */}
            <Typography variant="subtitle2" sx={{ mb: 1 }}>Доступ к аналитикам (подразделения)</Typography>

            <Box sx={{ fontSize: 13 }}>
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.5, borderBottom: '1px solid #e0e0e0', fontWeight: 600, color: '#666' }}>
                <Box sx={{ flex: 1 }}>Модель / Аналитика / Подразделение</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>Просмотр</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>Ввод</Box>
              </Box>

              {analyticPerms.map(model => {
                const mKey = `ap-${model.id}`
                const mOpen = expanded.has(mKey)
                const toggleM = () => setExpanded(prev => {
                  const next = new Set(prev)
                  if (next.has(mKey)) next.delete(mKey); else next.add(mKey)
                  return next
                })
                return (
                  <Box key={mKey}>
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.5, bgcolor: '#fafafa', cursor: 'pointer', '&:hover': { bgcolor: '#f0f0f0' } }}
                      onClick={toggleM}>
                      {mOpen ? <ExpandMoreOutlined sx={{ fontSize: 18, opacity: 0.5 }} /> : <ChevronRightOutlined sx={{ fontSize: 18, opacity: 0.5 }} />}
                      <Box sx={{ flex: 1, fontWeight: 600 }}>{model.name}</Box>
                    </Box>
                    {mOpen && model.analytics.map((analytic: any) => {
                      const aKey = `ap-a-${analytic.id}`
                      const aOpen = expanded.has(aKey)
                      const toggleA = () => setExpanded(prev => {
                        const next = new Set(prev)
                        if (next.has(aKey)) next.delete(aKey); else next.add(aKey)
                        return next
                      })
                      return (
                        <Box key={aKey}>
                          <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.25, pl: 4, cursor: 'pointer', '&:hover': { bgcolor: '#f8f8f8' } }}
                            onClick={toggleA}>
                            {aOpen ? <ExpandMoreOutlined sx={{ fontSize: 16, opacity: 0.4 }} /> : <ChevronRightOutlined sx={{ fontSize: 16, opacity: 0.4 }} />}
                            <Box sx={{ flex: 1, fontWeight: 500, color: '#555' }}>{analytic.name}</Box>
                          </Box>
                          {aOpen && analytic.records.map((rec: any) => (
                            <Box key={rec.id} sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.25, pl: 7, '&:hover': { bgcolor: '#f8f8f8' } }}>
                              <Box sx={{ flex: 1 }}>{rec.name}</Box>
                              <Box sx={{ width: 80, textAlign: 'center' }}>
                                <Checkbox size="small" checked={rec.can_view}
                                  onChange={e => {
                                    api.setAnalyticPermission({
                                      user_id: selectedId!, analytic_id: analytic.id,
                                      record_id: rec.id, can_view: e.target.checked, can_edit: rec.can_edit,
                                    }).then(loadPerms)
                                  }}
                                />
                              </Box>
                              <Box sx={{ width: 80, textAlign: 'center' }}>
                                <Checkbox size="small" checked={rec.can_edit}
                                  onChange={e => {
                                    api.setAnalyticPermission({
                                      user_id: selectedId!, analytic_id: analytic.id,
                                      record_id: rec.id, can_view: rec.can_view || e.target.checked, can_edit: e.target.checked,
                                    }).then(loadPerms)
                                  }}
                                />
                              </Box>
                            </Box>
                          ))}
                        </Box>
                      )
                    })}
                  </Box>
                )
              })}
            </Box>
          </Box>
        ) : (
          <Box sx={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999' }}>
            Выберите пользователя
          </Box>
        )}
      </Box>
    </Box>
  )
}
