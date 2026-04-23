import { useState, useEffect, useCallback } from 'react'
import {
  Box, Typography, IconButton, TextField, Tooltip, Checkbox,
  FormControlLabel, Switch, Divider,
  Dialog, DialogTitle, DialogContent, DialogActions, Button,
} from '@mui/material'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import PersonAddOutlined from '@mui/icons-material/PersonAddOutlined'
import PersonOutlined from '@mui/icons-material/PersonOutlined'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import ChevronRightOutlined from '@mui/icons-material/ChevronRightOutlined'
import DescriptionOutlined from '@mui/icons-material/DescriptionOutlined'
import FolderOutlined from '@mui/icons-material/FolderOutlined'
import CategoryOutlined from '@mui/icons-material/CategoryOutlined'
import KeyOutlined from '@mui/icons-material/KeyOutlined'
import { useTranslation } from 'react-i18next'
import * as api from '../api'

interface Props {
  open: boolean
  onClose: () => void
}

interface PermModel {
  id: string; name: string
  sheets: { id: string; name: string; can_view: boolean; can_edit: boolean }[]
  analytics?: { id: string; name: string; records: any[] }[]
}

export default function UsersDialog({ open, onClose }: Props) {
  const { t } = useTranslation()
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
    api.getAllPermissions(selectedId).then(data => {
      setPerms(data)
      setAnalyticPerms([]) // unified in perms now
    })
  }, [selectedId])

  const [pwDialogOpen, setPwDialogOpen] = useState(false)
  const [newPassword, setNewPassword] = useState('')

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
    const u = await api.createUser(t('users.newUser'))
    await loadUsers()
    setSelectedId(u.id)
  }

  const handleDelete = async () => {
    if (!selectedId) return
    if (!confirm(t('users.deleteConfirm'))) return
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
    if (!selectedId || !newPassword) return
    await api.resetPassword(selectedId, newPassword)
    setPwDialogOpen(false)
    setNewPassword('')
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
    const model = perms.find((m: any) => m.id === modelId)
    if (!model) return
    // Sheets
    for (const s of model.sheets) {
      const canView = field === 'can_view' ? value : s.can_view
      const canEdit = field === 'can_edit' ? value : s.can_edit
      await api.setSheetPermission(s.id, { user_id: selectedId, can_view: canView, can_edit: canEdit })
    }
    // Analytics — all records in all analytics
    const flat = (recs: any[]): any[] => recs.flatMap((r: any) => [r, ...flat(r.children || [])])
    if (model.analytics) {
      for (const a of model.analytics) {
        for (const r of flat(a.records)) {
          await api.setAnalyticPermission({
            user_id: selectedId, analytic_id: a.id, record_id: r.id,
            can_view: field === 'can_view' ? value : r.can_view || value,
            can_edit: field === 'can_edit' ? value : r.can_edit,
          })
        }
      }
    }
    loadPerms()
  }

  const modelChecked = (modelId: string, field: 'can_view' | 'can_edit') => {
    const model = perms.find((m: any) => m.id === modelId)
    if (!model) return false
    const flat = (recs: any[]): any[] => recs.flatMap((r: any) => [r, ...flat(r.children || [])])
    const allItems = [...model.sheets, ...(model.analytics || []).flatMap((a: any) => flat(a.records))]
    return allItems.length > 0 && allItems.every((s: any) => s[field])
  }

  const modelIndeterminate = (modelId: string, field: 'can_view' | 'can_edit') => {
    const model = perms.find((m: any) => m.id === modelId)
    if (!model) return false
    const flat = (recs: any[]): any[] => recs.flatMap((r: any) => [r, ...flat(r.children || [])])
    const allItems = [...model.sheets, ...(model.analytics || []).flatMap((a: any) => flat(a.records))]
    const on = allItems.filter((s: any) => s[field]).length
    return on > 0 && on < allItems.length
  }

  return (<>
    <Box sx={{ position: 'fixed', inset: 0, zIndex: 1300, bgcolor: '#fff', display: 'flex', flexDirection: 'column' }}>
      {/* Header */}
      <Box sx={{ display: 'flex', alignItems: 'center', px: 2, py: 1, borderBottom: '1px solid #e0e0e0' }}>
        <Typography variant="h6" sx={{ flex: 1 }}>{t('users.title')}</Typography>
        <IconButton size="small" onClick={onClose}><CloseOutlined fontSize="small" /></IconButton>
      </Box>

      <Box sx={{ flex: 1, display: 'flex', overflow: 'hidden' }}>
        {/* ─── Left: user list ─── */}
        <Box sx={{ width: 220, borderRight: '1px solid #e0e0e0', display: 'flex', flexDirection: 'column' }}>
          <Box sx={{ p: 1, borderBottom: '1px solid #f0f0f0', display: 'flex', gap: 0.5 }}>
            <Tooltip title={t('users.addUser')}>
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
                {!!u.can_admin && <Typography sx={{ fontSize: 10, color: '#1976d2', ml: 'auto' }}>{t('users.admin')}</Typography>}
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
                label={t('users.name')} size="small" value={username}
                onChange={e => setUsername(e.target.value)}
                onBlur={handleSaveName}
                onKeyDown={e => { if (e.key === 'Enter') handleSaveName() }}
                sx={{ flex: 1, maxWidth: 300 }}
              />
              <FormControlLabel
                control={<Switch checked={isAdmin} onChange={e => handleToggleAdmin(e.target.checked)} size="small" />}
                label={t('users.adminLabel')}
              />
              <Tooltip title={t('users.changePassword')}>
                <IconButton size="small" onClick={() => { setNewPassword(''); setPwDialogOpen(true) }}>
                  <KeyOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
              <Tooltip title={t('users.deleteUser')}>
                <IconButton size="small" color="error" onClick={handleDelete}>
                  <DeleteOutlineOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
            </Box>
            <Typography variant="caption" color="textSecondary" sx={{ mb: 3, display: 'block' }}>
              {t('users.created')} {createdAt}
            </Typography>

            <Divider sx={{ mb: 2 }} />

            {/* Unified permissions tree */}
            <Typography variant="subtitle2" sx={{ mb: 1 }}>{t('users.access')}</Typography>

            <Box sx={{ fontSize: 13 }}>
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.5, borderBottom: '1px solid #e0e0e0', fontWeight: 600, color: '#666' }}>
                <Box sx={{ flex: 1 }}>{t('users.modelSheetAnalytic')}</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>{t('users.view')}</Box>
                <Box sx={{ width: 80, textAlign: 'center' }}>{t('users.edit')}</Box>
              </Box>

              {perms.map((model: any) => {
                const isOpen = expanded.has(model.id)
                const toggle = () => setExpanded(prev => { const n = new Set(prev); n.has(model.id) ? n.delete(model.id) : n.add(model.id); return n })

                // Analytic record helpers
                const handleRecordPerm = (analyticId: string, recordId: string, field: string, value: boolean, rec: any) => {
                  const data = {
                    user_id: selectedId!, analytic_id: analyticId, record_id: recordId,
                    can_view: field === 'can_view' ? value : rec.can_view || value,
                    can_edit: field === 'can_edit' ? value : rec.can_edit,
                  }
                  api.setAnalyticPermission(data).then(loadPerms)
                }
                const setRecordTreePerm = (analyticId: string, records: any[], field: string, value: boolean) => {
                  const flat = (recs: any[]): any[] => recs.flatMap(r => [r, ...flat(r.children || [])])
                  Promise.all(flat(records).map(r => api.setAnalyticPermission({
                    user_id: selectedId!, analytic_id: analyticId, record_id: r.id,
                    can_view: field === 'can_view' ? value : r.can_view || value,
                    can_edit: field === 'can_edit' ? value : r.can_edit,
                  }))).then(loadPerms)
                }

                // Helpers to compute checked/indeterminate for analytics tree
                const flatAll = (recs: any[]): any[] => recs.flatMap(r => [r, ...flatAll(r.children || [])])
                const allAnalyticRecords = (model.analytics || []).flatMap((a: any) => flatAll(a.records))
                const analyticsFolderChecked = (field: string) => allAnalyticRecords.length > 0 && allAnalyticRecords.every((r: any) => r[field])
                const analyticsFolderIndeterminate = (field: string) => allAnalyticRecords.some((r: any) => r[field]) && !allAnalyticRecords.every((r: any) => r[field])
                const analyticChecked = (analytic: any, field: string) => { const flat = flatAll(analytic.records); return flat.length > 0 && flat.every((r: any) => r[field]) }
                const analyticIndeterminate = (analytic: any, field: string) => { const flat = flatAll(analytic.records); return flat.some((r: any) => r[field]) && !flat.every((r: any) => r[field]) }

                const renderRecords = (analyticId: string, records: any[], depth: number): any => {
                  return records.map((rec: any) => {
                    const hasChildren = rec.children && rec.children.length > 0
                    const rKey = `rec-${rec.id}`
                    const rOpen = expanded.has(rKey)
                    const toggleR = () => setExpanded(prev => { const n = new Set(prev); n.has(rKey) ? n.delete(rKey) : n.add(rKey); return n })
                    return (
                      <Box key={rec.id}>
                        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.25, pl: depth, '&:hover': { bgcolor: '#f8f8f8' }, cursor: hasChildren ? 'pointer' : 'default' }}
                          onClick={hasChildren ? toggleR : undefined}>
                          {hasChildren
                            ? (rOpen ? <ExpandMoreOutlined sx={{ fontSize: 14, opacity: 0.4 }} /> : <ChevronRightOutlined sx={{ fontSize: 14, opacity: 0.4 }} />)
                            : <Box sx={{ width: 18 }} />}
                          <Box sx={{ flex: 1, fontWeight: hasChildren ? 500 : 400 }}>{rec.name}</Box>
                          <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                            <Checkbox size="small" checked={rec.can_view}
                              onChange={e => {
                                if (hasChildren) setRecordTreePerm(analyticId, [rec], 'can_view', e.target.checked)
                                else handleRecordPerm(analyticId, rec.id, 'can_view', e.target.checked, rec)
                              }} />
                          </Box>
                          <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                            <Checkbox size="small" checked={rec.can_edit}
                              onChange={e => {
                                if (hasChildren) setRecordTreePerm(analyticId, [rec], 'can_edit', e.target.checked)
                                else handleRecordPerm(analyticId, rec.id, 'can_edit', e.target.checked, rec)
                              }} />
                          </Box>
                        </Box>
                        {hasChildren && rOpen && renderRecords(analyticId, rec.children, depth + 3)}
                      </Box>
                    )
                  })
                }

                return (
                <Box key={model.id}>
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.5, bgcolor: '#fafafa', '&:hover': { bgcolor: '#f0f0f0' }, cursor: 'pointer' }}
                    onClick={toggle}>
                    {isOpen ? <ExpandMoreOutlined sx={{ fontSize: 18, opacity: 0.5 }} /> : <ChevronRightOutlined sx={{ fontSize: 18, opacity: 0.5 }} />}
                    <Box sx={{ flex: 1, fontWeight: 600 }}>{model.name}</Box>
                    <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                      <Checkbox size="small"
                        checked={modelChecked(model.id, 'can_view')}
                        indeterminate={modelIndeterminate(model.id, 'can_view')}
                        onChange={e => handleModelPerm(model.id, 'can_view', e.target.checked)} />
                    </Box>
                    <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                      <Checkbox size="small"
                        checked={modelChecked(model.id, 'can_edit')}
                        indeterminate={modelIndeterminate(model.id, 'can_edit')}
                        onChange={e => handleModelPerm(model.id, 'can_edit', e.target.checked)} />
                    </Box>
                  </Box>

                  {isOpen && <>
                    {/* Sheets folder */}
                    {model.sheets?.length > 0 && (() => {
                      const sfKey = `sf-${model.id}`
                      const sfOpen = expanded.has(sfKey)
                      return (
                        <Box>
                          <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.25, pl: 4, cursor: 'pointer', '&:hover': { bgcolor: '#f8f8f8' } }}
                            onClick={() => setExpanded(prev => { const n = new Set(prev); n.has(sfKey) ? n.delete(sfKey) : n.add(sfKey); return n })}>
                            {sfOpen ? <ExpandMoreOutlined sx={{ fontSize: 16, opacity: 0.4 }} /> : <ChevronRightOutlined sx={{ fontSize: 16, opacity: 0.4 }} />}
                            <FolderOutlined sx={{ fontSize: 14, opacity: 0.4 }} />
                            <Box sx={{ flex: 1, fontWeight: 500, color: '#555' }}>{t('users.sheets')}</Box>
                            <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                              <Checkbox size="small"
                                checked={model.sheets.every((s: any) => s.can_view)}
                                indeterminate={model.sheets.some((s: any) => s.can_view) && !model.sheets.every((s: any) => s.can_view)}
                                onChange={e => { model.sheets.forEach((s: any) => handleSheetPerm(s.id, 'can_view', e.target.checked)) }} />
                            </Box>
                            <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                              <Checkbox size="small"
                                checked={model.sheets.every((s: any) => s.can_edit)}
                                indeterminate={model.sheets.some((s: any) => s.can_edit) && !model.sheets.every((s: any) => s.can_edit)}
                                onChange={e => { model.sheets.forEach((s: any) => handleSheetPerm(s.id, 'can_edit', e.target.checked)) }} />
                            </Box>
                          </Box>
                          {sfOpen && model.sheets.map((sheet: any) => (
                            <Box key={sheet.id} sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1, py: 0.25, pl: 7, '&:hover': { bgcolor: '#f8f8f8' } }}>
                              <DescriptionOutlined sx={{ fontSize: 14, opacity: 0.4 }} />
                              <Box sx={{ flex: 1 }}>{sheet.name}</Box>
                              <Box sx={{ width: 80, textAlign: 'center' }}>
                                <Checkbox size="small" checked={sheet.can_view}
                                  onChange={e => handleSheetPerm(sheet.id, 'can_view', e.target.checked)} />
                              </Box>
                              <Box sx={{ width: 80, textAlign: 'center' }}>
                                <Checkbox size="small" checked={sheet.can_edit}
                                  onChange={e => handleSheetPerm(sheet.id, 'can_edit', e.target.checked)} />
                              </Box>
                            </Box>
                          ))}
                        </Box>
                      )
                    })()}

                    {/* Analytics folder */}
                    {model.analytics?.length > 0 && (() => {
                      const afKey = `af-${model.id}`
                      const afOpen = expanded.has(afKey)
                      return (
                        <Box>
                          <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.25, pl: 4, cursor: 'pointer', '&:hover': { bgcolor: '#f8f8f8' } }}
                            onClick={() => setExpanded(prev => { const n = new Set(prev); n.has(afKey) ? n.delete(afKey) : n.add(afKey); return n })}>
                            {afOpen ? <ExpandMoreOutlined sx={{ fontSize: 16, opacity: 0.4 }} /> : <ChevronRightOutlined sx={{ fontSize: 16, opacity: 0.4 }} />}
                            <FolderOutlined sx={{ fontSize: 14, opacity: 0.4 }} />
                            <Box sx={{ flex: 1, fontWeight: 500, color: '#555' }}>{t('users.analytics')}</Box>
                            <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                              <Checkbox size="small"
                                checked={analyticsFolderChecked('can_view')}
                                indeterminate={analyticsFolderIndeterminate('can_view')}
                                onChange={e => { model.analytics.forEach((a: any) => setRecordTreePerm(a.id, a.records, 'can_view', e.target.checked)) }} />
                            </Box>
                            <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                              <Checkbox size="small"
                                checked={analyticsFolderChecked('can_edit')}
                                indeterminate={analyticsFolderIndeterminate('can_edit')}
                                onChange={e => { model.analytics.forEach((a: any) => setRecordTreePerm(a.id, a.records, 'can_edit', e.target.checked)) }} />
                            </Box>
                          </Box>
                          {afOpen && model.analytics.map((analytic: any) => {
                            const aKey = `a-${analytic.id}`
                            const aOpen = expanded.has(aKey)
                            return (
                              <Box key={analytic.id}>
                                <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, px: 1, py: 0.25, pl: 7, cursor: 'pointer', '&:hover': { bgcolor: '#f8f8f8' } }}
                                  onClick={() => setExpanded(prev => { const n = new Set(prev); n.has(aKey) ? n.delete(aKey) : n.add(aKey); return n })}>
                                  {aOpen ? <ExpandMoreOutlined sx={{ fontSize: 14, opacity: 0.4 }} /> : <ChevronRightOutlined sx={{ fontSize: 14, opacity: 0.4 }} />}
                                  <CategoryOutlined sx={{ fontSize: 14, opacity: 0.4 }} />
                                  <Box sx={{ flex: 1, fontWeight: 500, color: '#555' }}>{analytic.name}</Box>
                                  <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                                    <Checkbox size="small"
                                      checked={analyticChecked(analytic, 'can_view')}
                                      indeterminate={analyticIndeterminate(analytic, 'can_view')}
                                      onChange={e => setRecordTreePerm(analytic.id, analytic.records, 'can_view', e.target.checked)} />
                                  </Box>
                                  <Box sx={{ width: 80, textAlign: 'center' }} onClick={e => e.stopPropagation()}>
                                    <Checkbox size="small"
                                      checked={analyticChecked(analytic, 'can_edit')}
                                      indeterminate={analyticIndeterminate(analytic, 'can_edit')}
                                      onChange={e => setRecordTreePerm(analytic.id, analytic.records, 'can_edit', e.target.checked)} />
                                  </Box>
                                </Box>
                                {aOpen && renderRecords(analytic.id, analytic.records, 10)}
                              </Box>
                            )
                          })}
                        </Box>
                      )
                    })()}
                  </>}
                </Box>
                )
              })}

              {perms.length === 0 && (
                <Typography sx={{ py: 2, color: '#999', textAlign: 'center' }}>{t('users.noModels')}</Typography>
              )}
            </Box>
          </Box>
        ) : (
          <Box sx={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999' }}>
            {t('users.selectUser')}
          </Box>
        )}
      </Box>
    </Box>

    <Dialog open={pwDialogOpen} onClose={() => setPwDialogOpen(false)} maxWidth="xs" fullWidth>
      <DialogTitle>{t('users.changePassword')}</DialogTitle>
      <DialogContent>
        <TextField
          label={t('users.newPassword')} fullWidth type="password"
          value={newPassword} onChange={e => setNewPassword(e.target.value)}
          autoComplete="new-password" name="new-password"
          sx={{ mt: 1 }} autoFocus
        />
      </DialogContent>
      <DialogActions>
        <Button onClick={() => setPwDialogOpen(false)}>{t('common.cancel')}</Button>
        <Button variant="contained" disabled={!newPassword} onClick={handleResetPassword}>{t('common.save')}</Button>
      </DialogActions>
    </Dialog>
  </>)
}
