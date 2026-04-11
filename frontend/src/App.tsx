import { useState, useCallback, useEffect, useRef } from 'react'
import {
  IconButton, Tooltip, Badge, Select, MenuItem, FormControl,
  Dialog, DialogTitle, DialogContent, DialogActions, Button, TextField, CircularProgress,
  ToggleButton, ToggleButtonGroup, Typography, Box,
} from '@mui/material'
import RefreshOutlined from '@mui/icons-material/RefreshOutlined'
import MenuOutlined from '@mui/icons-material/MenuOutlined'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import SettingsOutlined from '@mui/icons-material/SettingsOutlined'
import TableChartOutlined from '@mui/icons-material/TableChartOutlined'
import FunctionsOutlined from '@mui/icons-material/FunctionsOutlined'
import PeopleOutlined from '@mui/icons-material/PeopleOutlined'
import FileUploadOutlined from '@mui/icons-material/FileUploadOutlined'
import type { TreeSelection } from './types'
import LeftPanel from './panels/LeftPanel'
import CenterPanel from './panels/CenterPanel'
import Splitter from './components/Splitter'
import UsersDialog from './components/UsersDialog'
import PivotGrid from './features/sheet/PivotGrid'
import { PendingProvider, usePending } from './store/PendingContext'
import * as api from './api'
import './App.css'

type AppMode = 'settings' | 'data' | 'formulas'

function SaveButton() {
  const { isDirty, flush } = usePending()
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); if (isDirty) flush() }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [isDirty, flush])
  return (
    <Tooltip title={isDirty ? 'Сохранить (Ctrl+S)' : 'Нет изменений'}>
      <span>
        <IconButton size="small" disabled={!isDirty} onClick={flush} sx={{ color: isDirty ? '#1976d2' : undefined }}>
          <Badge variant="dot" color="error" invisible={!isDirty}><SaveOutlined fontSize="small" /></Badge>
        </IconButton>
      </span>
    </Tooltip>
  )
}

function ImportDialog({ open, onClose, onImported }: {
  open: boolean; onClose: () => void; onImported: (modelId: string) => void
}) {
  const [file, setFile] = useState<File | null>(null)
  const [modelName, setModelName] = useState('')
  const [loading, setLoading] = useState(false)
  const [log, setLog] = useState<string[]>([])
  const [done, setDone] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)
  const logRef = useRef<HTMLDivElement>(null)

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0]
    if (f) {
      setFile(f)
      if (!modelName) setModelName(f.name.replace(/\.xlsx?$/i, ''))
    }
  }

  const handleImport = async () => {
    if (!file || !modelName) return
    setLoading(true)
    setLog([])
    setDone(false)
    try {
      const result = await api.importExcelModelStream(file, modelName, (msg, data) => {
        setLog(prev => [...prev, msg])
        setTimeout(() => logRef.current?.scrollTo(0, logRef.current.scrollHeight), 50)
        if (data?.done) {
          setDone(true)
          onImported(data.model_id)
        }
      })
    } catch (err) {
      setLog(prev => [...prev, `❌ Ошибка: ${(err as Error).message}`])
    } finally {
      setLoading(false)
    }
  }

  const handleClose = () => {
    if (loading) return
    onClose()
    setFile(null)
    setModelName('')
    setLog([])
    setDone(false)
  }

  return (
    <Dialog open={open} onClose={handleClose} maxWidth="sm" fullWidth>
      <DialogTitle>Импорт модели из Excel</DialogTitle>
      <DialogContent sx={{ display: 'flex', flexDirection: 'column', gap: 2, pt: 1 }}>
        <input
          ref={fileRef} type="file" accept=".xlsx,.xls"
          style={{ display: 'none' }} onChange={handleFileChange}
        />
        <Button variant="outlined" onClick={() => fileRef.current?.click()} disabled={loading}>
          {file ? file.name : 'Выбрать файл (.xlsx)'}
        </Button>
        <TextField
          label="Название модели" value={modelName}
          onChange={e => setModelName(e.target.value)}
          fullWidth size="small" disabled={loading}
        />
        {log.length > 0 && (
          <Box
            ref={logRef}
            sx={{
              maxHeight: 260, overflow: 'auto', bgcolor: '#f8f9fa', borderRadius: 1,
              p: 1.5, fontFamily: 'monospace', fontSize: 12, lineHeight: 1.6,
              border: '1px solid #e0e0e0',
            }}
          >
            {log.map((line, i) => (
              <div key={i} style={{ color: line.startsWith('❌') ? '#c62828' : line.startsWith('✅') ? '#2e7d32' : '#333' }}>
                {line}
              </div>
            ))}
            {loading && !done && (
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mt: 0.5, color: '#1976d2' }}>
                <CircularProgress size={12} /> <span>работаю...</span>
              </Box>
            )}
          </Box>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={handleClose} disabled={loading}>
          {done ? 'Закрыть' : 'Отмена'}
        </Button>
        {!done && (
          <Button
            variant="contained"
            disabled={!file || !modelName || loading}
            onClick={handleImport}
            startIcon={loading ? <CircularProgress size={16} /> : undefined}
          >
            {loading ? 'Импорт...' : 'Импортировать'}
          </Button>
        )}
      </DialogActions>
    </Dialog>
  )
}

function AppInner() {
  const [mode, setMode] = useState<AppMode>('settings')
  const [selection, setSelection] = useState<TreeSelection | null>(null)
  const [leftWidth, setLeftWidth] = useState(280)
  const [leftOpen, setLeftOpen] = useState(true)
  const [refreshKey, setRefreshKey] = useState(0)
  const [showUsers, setShowUsers] = useState(false)
  const [showImport, setShowImport] = useState(false)
  const [expandAfterCreate, setExpandAfterCreate] = useState<any>(null)
  const [users, setUsers] = useState<any[]>([])
  const [currentUserId, setCurrentUserId] = useState('')

  useEffect(() => {
    api.listUsers().then(u => {
      setUsers(u)
      if (u.length > 0) {
        // Set first user if current is empty or no longer exists
        const ids = new Set(u.map((x: any) => x.id))
        setCurrentUserId(prev => (prev && ids.has(prev)) ? prev : u[0].id)
      }
    })
  }, [showUsers])

  const onRefresh = useCallback(() => setRefreshKey(k => k + 1), [])

  const onCreated = useCallback((info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => {
    setExpandAfterCreate({ modelId: info.modelId, folder: info.folder, selectId: info.id, selectType: info.type })
    setSelection({ type: info.type, id: info.id, modelId: info.modelId })
    setRefreshKey(k => k + 1)
  }, [])

  const handleImported = useCallback((modelId: string) => {
    setSelection({ type: 'model', id: modelId, modelId })
    setRefreshKey(k => k + 1)
  }, [])

  // When switching to data/formulas mode, if a sheet is selected — keep it.
  // When clicking a sheet in left panel — select it and switch to data mode if needed.
  const handleSelect = useCallback((sel: TreeSelection | null) => {
    setSelection(sel)
  }, [])

  const isSheetSelected = selection?.type === 'sheet'
  const isDataMode = mode === 'data' || mode === 'formulas'
  const currentUser = users.find(u => u.id === currentUserId)
  const isAdmin = !!currentUser?.can_admin

  // Non-admin users can only use data mode
  useEffect(() => {
    if (currentUserId && !isAdmin && mode !== 'data') {
      setMode('data')
    }
  }, [currentUserId, isAdmin])

  return (
    <PendingProvider onFlushed={onRefresh}>
      <div className="app-root">
        <div className="app-toolbar">
          <Tooltip title={leftOpen ? "Скрыть панель" : "Показать панель"}>
            <IconButton size="small" onClick={() => setLeftOpen(v => !v)}>
              <MenuOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
          <Tooltip title="Обновить">
            <IconButton size="small" onClick={onRefresh}>
              <RefreshOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
          <SaveButton />

          {/* Mode toggle */}
          <ToggleButtonGroup
            size="small" exclusive
            value={mode}
            onChange={(_, v) => { if (v) setMode(v) }}
            sx={{ '& .MuiToggleButton-root': { py: 0.25, px: 1, fontSize: 12, textTransform: 'none' } }}
          >
            {isAdmin && (
              <ToggleButton value="settings">
                <Tooltip title="Настройки модели"><SettingsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
            <ToggleButton value="data">
              <Tooltip title="Просмотр / ввод данных"><TableChartOutlined sx={{ fontSize: 16 }} /></Tooltip>
            </ToggleButton>
            {isAdmin && (
              <ToggleButton value="formulas">
                <Tooltip title="Формулы и правила"><FunctionsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
          </ToggleButtonGroup>

          {isAdmin && (
            <Tooltip title="Импорт модели из Excel">
              <IconButton size="small" onClick={() => setShowImport(true)}>
                <FileUploadOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}

          <div style={{ flex: 1 }} />

          {/* User selector — right aligned */}
          {users.length > 0 && (
            <FormControl size="small" sx={{ minWidth: 120 }}>
              <Select
                value={currentUserId}
                onChange={e => setCurrentUserId(e.target.value)}
                variant="standard"
                disableUnderline
                sx={{ fontSize: 12 }}
              >
                {users.map(u => <MenuItem key={u.id} value={u.id} sx={{ fontSize: 12 }}>{u.username}</MenuItem>)}
              </Select>
            </FormControl>
          )}

          <Tooltip title="Пользователи">
            <IconButton size="small" onClick={() => setShowUsers(true)}>
              <PeopleOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
        </div>

        <div className="app-body">
          {leftOpen && <>
            <div style={{ width: leftWidth, minWidth: 180, flexShrink: 0 }}>
              <LeftPanel
                selection={selection} onSelect={handleSelect}
                refreshKey={refreshKey} expandAfterCreate={expandAfterCreate} onCreated={onCreated}
                sheetsOnly={isDataMode} currentUserId={isDataMode ? currentUserId : undefined}
              />
            </div>
            <Splitter onResize={d => setLeftWidth(w => Math.max(180, w + d))} />
          </>}

          {/* Center area: settings or pivot grid */}
          {mode === 'settings' ? (
            <CenterPanel selection={selection} onRefresh={onRefresh} />
          ) : isSheetSelected ? (
            <PivotGrid
              key={`${selection.id}-${refreshKey}`}
              sheetId={selection.id} modelId={selection.modelId}
              currentUserId={currentUserId}
              mode={mode === 'formulas' ? 'settings' : 'data'}
            />
          ) : (
            <div className="panel-center" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999' }}>
              Выберите лист для просмотра данных
            </div>
          )}
        </div>

        <UsersDialog open={showUsers} onClose={() => setShowUsers(false)} />
        <ImportDialog open={showImport} onClose={() => setShowImport(false)} onImported={handleImported} />
      </div>
    </PendingProvider>
  )
}

export default function App() { return <AppInner /> }
