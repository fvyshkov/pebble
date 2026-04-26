import { useState, useCallback, useEffect, useRef } from 'react'
import { useTranslation } from 'react-i18next'
import {
  IconButton, Tooltip, Badge, Select, MenuItem, FormControl,
  Dialog, DialogTitle, DialogContent, DialogActions, Button, TextField, CircularProgress,
  ToggleButton, ToggleButtonGroup, Typography, Box, Chip, Popover, List, ListItem, ListItemIcon, ListItemText,
} from '@mui/material'
import RefreshOutlined from '@mui/icons-material/RefreshOutlined'
import MenuOutlined from '@mui/icons-material/MenuOutlined'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import SettingsOutlined from '@mui/icons-material/SettingsOutlined'
import TableChartOutlined from '@mui/icons-material/TableChartOutlined'
import FunctionsOutlined from '@mui/icons-material/FunctionsOutlined'
import PeopleOutlined from '@mui/icons-material/PeopleOutlined'
import LogoutOutlined from '@mui/icons-material/LogoutOutlined'
import CalculateOutlined from '@mui/icons-material/CalculateOutlined'
import FileUploadOutlined from '@mui/icons-material/FileUploadOutlined'
import SmartToyOutlined from '@mui/icons-material/SmartToyOutlined'
import CheckCircleOutlined from '@mui/icons-material/CheckCircleOutlined'
import ErrorOutlineOutlined from '@mui/icons-material/ErrorOutlineOutlined'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import WarningAmberOutlined from '@mui/icons-material/WarningAmberOutlined'
import DeleteSweepOutlined from '@mui/icons-material/DeleteSweepOutlined'
import TranslateOutlined from '@mui/icons-material/TranslateOutlined'
import UndoOutlined from '@mui/icons-material/UndoOutlined'
import ArrowDropDownOutlined from '@mui/icons-material/ArrowDropDownOutlined'
import FormatListNumberedOutlined from '@mui/icons-material/FormatListNumberedOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import type { TreeSelection } from './types'
import LoginPage from './features/auth/LoginPage'
import LeftPanel from './panels/LeftPanel'
import CenterPanel from './panels/CenterPanel'
import Splitter from './components/Splitter'
import UsersDialog from './components/UsersDialog'
// PivotGrid removed — AG Grid is the only grid
import PivotGridAG, { type PivotGridAGHandle } from './features/sheet/PivotGridAG'
import ChatPanel from './features/chat/ChatPanel'
import ChartPanel, { type ChartConfig } from './features/chart/ChartPanel'
import PresentationPanel from './features/presentation/PresentationPanel'
import { PendingProvider, usePending } from './store/PendingContext'
import { LANGUAGES, changeLanguage, currentLang } from './i18n'
import * as api from './api'
import './App.css'

type AppMode = 'settings' | 'data' | 'formulas'

function SaveButton() {
  const { isDirty, flush } = usePending()
  const { t } = useTranslation()
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); if (isDirty) flush() }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [isDirty, flush])
  return (
    <Tooltip title={isDirty ? t('app.save') : t('app.noChanges')}>
      <span>
        <IconButton size="small" disabled={!isDirty} onClick={flush} sx={{ color: isDirty ? '#1976d2' : undefined }}>
          <Badge variant="dot" color="error" invisible={!isDirty}><SaveOutlined fontSize="small" /></Badge>
        </IconButton>
      </span>
    </Tooltip>
  )
}

function LanguageSwitcher() {
  const [lang, setLang] = useState(currentLang())
  return (
    <Select
      size="small"
      value={lang}
      onChange={e => { const v = e.target.value; changeLanguage(v); setLang(v) }}
      sx={{ fontSize: 12, height: 28, minWidth: 56, '& .MuiSelect-select': { py: 0.25, px: 1 } }}
    >
      {LANGUAGES.map(l => (
        <MenuItem key={l.code} value={l.code} sx={{ fontSize: 12 }}>{l.label}</MenuItem>
      ))}
    </Select>
  )
}

function ImportDialog({ open, onClose, onImported, initialFile }: {
  open: boolean; onClose: () => void; onImported: (modelId: string) => void; initialFile?: File | null
}) {
  const { t } = useTranslation()
  const [file, setFile] = useState<File | null>(null)
  const [modelName, setModelName] = useState('')

  // When initialFile is provided, pre-fill file and model name
  useEffect(() => {
    if (initialFile && open) {
      setFile(initialFile)
      setModelName(initialFile.name.replace(/\.xlsx?$/i, ''))
    }
  }, [initialFile, open])
  const [loading, setLoading] = useState(false)
  const [log, setLog] = useState<string[]>([])
  const [done, setDone] = useState(false)
  const [elapsed, setElapsed] = useState(0)
  const [progress, setProgress] = useState({ current: 0, total: 0 })
  const elapsedRef = useRef<ReturnType<typeof setInterval>>()
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
    setElapsed(0)
    elapsedRef.current = setInterval(() => setElapsed(t => t + 1), 1000)
    setProgress({ current: 0, total: 0 })
    try {
      const result = await api.importExcelModelStream(file, modelName, (msg, data) => {
        setLog(prev => [...prev, msg])
        setTimeout(() => logRef.current?.scrollTo(0, logRef.current.scrollHeight), 50)
        // Parse progress from messages like "(3/7)"
        const pm = msg.match(/\((\d+)\/(\d+)\)/)
        if (pm) setProgress({ current: parseInt(pm[1]), total: parseInt(pm[2]) })
        if (data?.done) {
          setDone(true)
          onImported(data.model_id)
        }
      }, currentLang())
    } catch (err) {
      setLog(prev => [...prev, `[ERR]${t('import.error')}: ${(err as Error).message}`])
    } finally {
      setLoading(false)
      clearInterval(elapsedRef.current)
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
      <DialogTitle>{t('import.title')}</DialogTitle>
      <DialogContent sx={{ display: 'flex', flexDirection: 'column', gap: 2, pt: 1 }}>
        <input
          ref={fileRef} type="file" accept=".xlsx,.xls"
          style={{ display: 'none' }} onChange={handleFileChange}
        />
        <Button variant="outlined" onClick={() => fileRef.current?.click()} disabled={loading}>
          {file ? file.name : t('import.selectFile')}
        </Button>
        <TextField
          label={t('import.modelName')} value={modelName}
          onChange={e => setModelName(e.target.value)}
          fullWidth size="small" disabled={loading}
        />
        {loading && !done && (
          <Box sx={{ width: '100%', mb: 1 }}>
            {progress.total > 0 ? (
              <>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#666', mb: 0.5 }}>
                  <span>{t('import.indicators')}: {progress.current} {t('import.of')} {progress.total}</span>
                  <span>{Math.round(progress.current / progress.total * 100)}%</span>
                </Box>
                <Box sx={{ width: '100%', height: 6, bgcolor: '#e0e0e0', borderRadius: 3 }}>
                  <Box sx={{ width: `${progress.current / progress.total * 100}%`, height: '100%', bgcolor: '#1976d2', borderRadius: 3, transition: 'width 0.3s' }} />
                </Box>
              </>
            ) : (
              <Box sx={{ width: '100%', height: 6, bgcolor: '#e0e0e0', borderRadius: 3, overflow: 'hidden' }}>
                <Box sx={{ width: '30%', height: '100%', bgcolor: '#1976d2', borderRadius: 3,
                  animation: 'indeterminate 1.5s ease-in-out infinite',
                  '@keyframes indeterminate': { '0%': { transform: 'translateX(-100%)' }, '100%': { transform: 'translateX(400%)' } },
                }} />
              </Box>
            )}
          </Box>
        )}
        {log.length > 0 && (
          <Box
            ref={logRef}
            sx={{
              maxHeight: 260, overflow: 'auto', bgcolor: '#f8f9fa', borderRadius: 1,
              p: 1.5, fontFamily: 'monospace', fontSize: 12, lineHeight: 1.6,
              border: '1px solid #e0e0e0',
            }}
          >
            {log.map((line, i) => {
              const isError = line.includes('[ERR]')
              const isSuccess = line.includes('[DONE]')
              const isWarn = line.includes('[WARN]')
              const hasCheck = line.includes('[OK]')
              const color = isError ? '#c62828' : isSuccess ? '#2e7d32' : isWarn ? '#e65100' : '#333'
              const text = line.replace(/\[(ERR|DONE|WARN|OK)\]/g, '').trim()
              const Icon = isError ? ErrorOutlineOutlined
                : isSuccess ? CheckCircleOutlined
                : isWarn ? WarningAmberOutlined
                : hasCheck ? CheckCircleOutlined
                : null
              return (
                <Box key={i} sx={{ display: 'flex', alignItems: 'flex-start', gap: 0.5, color }}>
                  {Icon && <Icon sx={{ fontSize: 14, mt: '2px', flexShrink: 0 }} />}
                  <span>{text}</span>
                </Box>
              )
            })}
            {loading && !done && (() => {
              const mins = Math.floor(elapsed / 60)
              const secs = String(elapsed % 60).padStart(2, '0')
              return (
                <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mt: 0.5, color: '#1976d2' }}>
                  <CircularProgress size={12} /> <span>{t('app.working')} {elapsed > 0 ? `${mins}:${secs}` : ''}</span>
                </Box>
              )
            })()}
          </Box>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={handleClose} disabled={loading} startIcon={done ? <CloseOutlined /> : undefined}>
          {done ? t('import.close') : t('app.cancel')}
        </Button>
        {!done && (
          <Button
            variant="contained"
            disabled={!file || !modelName || loading}
            onClick={handleImport}
            startIcon={loading ? <CircularProgress size={16} /> : undefined}
          >
            {loading ? t('import.importing') : t('import.importBtn')}
          </Button>
        )}
      </DialogActions>
    </Dialog>
  )
}

function AppInner({ authUser, onLogout }: { authUser?: { id: string; username: string; can_admin: boolean }; onLogout?: () => void }) {
  const { t } = useTranslation()
  const [mode, setMode] = useState<AppMode>('data')
  const [selection, setSelection] = useState<TreeSelection | null>(null)
  const [leftWidth, setLeftWidth] = useState(280)
  const [leftOpen, setLeftOpen] = useState(true)
  const [refreshKey, setRefreshKey] = useState(0)
  const [showUsers, setShowUsers] = useState(false)
  const [showImport, setShowImport] = useState(false)
  const [confirmDeleteAll, setConfirmDeleteAll] = useState(false)
  const [deletingAll, setDeletingAll] = useState(false)
  const [expandAfterCreate, setExpandAfterCreate] = useState<any>(null)
  const [users, setUsers] = useState<any[]>([])
  const [currentUserId, setCurrentUserId] = useState(authUser?.id || '')
  const [calcMode, setCalcMode] = useState<'auto' | 'manual'>(() =>
    (localStorage.getItem('pebble_calcMode') as 'auto' | 'manual') || 'auto'
  )
  const [calcRunning, setCalcRunning] = useState(false)
  const [calcProgress, setCalcProgress] = useState<{
    done: number; total: number; sheet?: string;
    computed?: number; totalCells?: number; startedAt?: number;
  } | null>(null)
  const [chatOpen, setChatOpen] = useState(false)
  const [chatWidth, setChatWidth] = useState<number>(() => {
    const v = parseInt(localStorage.getItem('pebble_chatWidth') || '', 10)
    return Number.isFinite(v) && v >= 280 ? v : 400
  })
  useEffect(() => { localStorage.setItem('pebble_chatWidth', String(chatWidth)) }, [chatWidth])
  const [chatImportFile, setChatImportFile] = useState<File | null>(null)
  // ── Undo state ──
  const gridRef = useRef<PivotGridAGHandle>(null)
  const [hasUndo, setHasUndo] = useState(false)
  const [undoAnchor, setUndoAnchor] = useState<HTMLElement | null>(null)
  const [undoItems, setUndoItems] = useState<any[]>([])
  const [undoHoverIdx, setUndoHoverIdx] = useState<number | null>(null)

  // ── Analytics order dialog ──
  const [reorderOpen, setReorderOpen] = useState(false)
  const [reorderItems, setReorderItems] = useState<string[]>([])
  const [reorderNames, setReorderNames] = useState<Record<string, string>>({})
  const reorderDragIdx = useRef<number | null>(null)
  const [reorderDragOver, setReorderDragOver] = useState<number | null>(null)

  // AG Grid is the only grid mode now.
  const [chartConfig, setChartConfig] = useState<ChartConfig | null>(null)
  const [presentation, setPresentation] = useState<{ html: string; title: string } | null>(null)
  const calcedModelsRef = useRef<Set<string>>(new Set())

  useEffect(() => {
    api.listUsers().then(u => {
      setUsers(u)
      // If authUser exists, use it; otherwise first user
      if (authUser) {
        setCurrentUserId(authUser.id)
      } else if (u.length > 0) {
        const ids = new Set(u.map((x: any) => x.id))
        setCurrentUserId(prev => (prev && ids.has(prev)) ? prev : u[0].id)
      }
    })
  }, [showUsers])

  // Persist calcMode to localStorage
  useEffect(() => { localStorage.setItem('pebble_calcMode', calcMode) }, [calcMode])

  const onRefresh = useCallback(() => setRefreshKey(k => k + 1), [])

  const handleDeleteAllModels = async () => {
    setDeletingAll(true)
    try {
      const models = await api.listModels()
      for (const m of models) {
        await api.deleteModel(m.id)
      }
      setConfirmDeleteAll(false)
      setSelection(null)
      onRefresh()
    } finally {
      setDeletingAll(false)
    }
  }

  // ── Global shortcuts ──────────────────────────────────────────────────
  const lastSpaceRef = useRef<number>(0)
  useEffect(() => {
    const isEditable = () => {
      const el = document.activeElement as HTMLElement | null
      if (!el) return false
      const tag = el.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return true
      if (el.isContentEditable) return true
      return false
    }
    const stripTrailingSpace = () => {
      const el = document.activeElement as (HTMLElement & { value?: string }) | null
      if (!el) return
      if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
        const v = (el as HTMLInputElement | HTMLTextAreaElement).value
        if (v && v.endsWith(' ')) {
          const setter = Object.getOwnPropertyDescriptor(
            Object.getPrototypeOf(el), 'value',
          )?.set
          setter?.call(el, v.slice(0, -1))
          el.dispatchEvent(new Event('input', { bubbles: true }))
        }
        return
      }
      if (el.isContentEditable) {
        const t = el.textContent || ''
        if (t.endsWith(' ')) {
          el.textContent = t.slice(0, -1)
          el.dispatchEvent(new Event('input', { bubbles: true }))
        }
      }
    }
    const onKey = async (ev: KeyboardEvent) => {
      if (ev.key === ' ' && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ev.shiftKey) {
        const now = Date.now()
        if (now - lastSpaceRef.current < 400) {
          ev.preventDefault()
          lastSpaceRef.current = 0
          if (isEditable()) stripTrailingSpace()
          setChatOpen(true)
          setTimeout(() => {
            window.dispatchEvent(new CustomEvent('pebble:toggleVoice'))
          }, 50)
        } else {
          lastSpaceRef.current = now
        }
        return
      }
      // Cmd/Ctrl+J → toggle AI chat panel
      if ((ev.metaKey || ev.ctrlKey) && !ev.shiftKey && !ev.altKey && ev.key.toLowerCase() === 'j') {
        ev.preventDefault()
        setChatOpen(v => !v)
        return
      }
      // Cmd/Ctrl+Z → history undo (point update, no full reload)
      if ((ev.metaKey || ev.ctrlKey) && !ev.shiftKey && ev.key.toLowerCase() === 'z') {
        if (isEditable()) return
        const modelId = selection?.modelId
        if (!modelId) return
        ev.preventDefault()
        try {
          const hist = await api.getModelHistory(modelId, 1)
          if (hist.length > 0) {
            const result = await api.undoChanges(modelId, hist[0].id)
            if (result.all_cells && gridRef.current) {
              gridRef.current.applyCellUpdates(result.all_cells)
            } else {
              setRefreshKey(k => k + 1)
            }
            // Refresh undo state
            api.getModelHistory(modelId, 1).then(h => setHasUndo(h.length > 0)).catch(() => {})
          }
        } catch { /* ignore */ }
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selection?.modelId])

  const onCreated = useCallback((info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => {
    setExpandAfterCreate({ modelId: info.modelId, folder: info.folder, selectId: info.id, selectType: info.type })
    setSelection({ type: info.type, id: info.id, modelId: info.modelId })
    setRefreshKey(k => k + 1)
  }, [])

  const handleImported = useCallback((modelId: string) => {
    setSelection({ type: 'model', id: modelId, modelId })
    setRefreshKey(k => k + 1)
    // Auto-recalc after import
    const startedAt = Date.now()
    setCalcRunning(true)
    setCalcProgress({ done: 0, total: 1, startedAt })
    api.calculateModelStream(modelId, (data) => {
      if (data.phase === 'start') {
        setCalcProgress({ done: 0, total: data.total_sheets || 1, totalCells: data.total_cells ?? undefined, computed: 0, startedAt })
      } else if (data.phase === 'sheet_done') {
        setCalcProgress({ done: data.done || 0, total: data.total_sheets || 1, sheet: data.sheet, totalCells: data.total_cells ?? undefined, computed: data.computed ?? undefined, startedAt })
      } else if (data.phase === 'done') {
        setCalcProgress(null); setCalcRunning(false)
      }
    }).catch(() => { setCalcRunning(false); setCalcProgress(null) })
  }, [])

  const handleSelect = useCallback((sel: TreeSelection | null) => {
    setSelection(sel)
  }, [])

  const isSheetSelected = selection?.type === 'sheet'
  const isDataMode = mode === 'data' || mode === 'formulas'
  const currentUser = users.find(u => u.id === currentUserId)
  const isAdmin = !!authUser?.can_admin || !!currentUser?.can_admin

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
          <Tooltip title={leftOpen ? t('app.hidePanel') : t('app.showPanel')}>
            <IconButton size="small" onClick={() => setLeftOpen(v => !v)}>
              <MenuOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
          {isAdmin && (
            <Tooltip title={t('app.importExcel')}>
              <IconButton size="small" onClick={() => setShowImport(true)}>
                <FileUploadOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}
          <Tooltip title={t('app.refresh')}>
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
            <ToggleButton value="data">
              <Tooltip title={t('app.dataMode')}><TableChartOutlined sx={{ fontSize: 16 }} /></Tooltip>
            </ToggleButton>
            {isAdmin && (
              <ToggleButton value="formulas">
                <Tooltip title={t('app.formulasMode')}><FunctionsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
            {isAdmin && (
              <ToggleButton value="settings">
                <Tooltip title={t('app.settingsMode')}><SettingsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
          </ToggleButtonGroup>


          {/* Calc mode toggle + calculate button */}
          <Tooltip title={calcMode === 'auto' ? t('app.autoCalc') : t('app.manualCalc')}>
            <Chip
              size="small"
              label={calcMode === 'auto' ? t('app.auto') : t('app.manual')}
              variant={calcMode === 'auto' ? 'filled' : 'outlined'}
              color={calcMode === 'auto' ? 'success' : 'default'}
              onClick={() => setCalcMode(prev => prev === 'auto' ? 'manual' : 'auto')}
              sx={{ fontSize: 11, cursor: 'pointer' }}
            />
          </Tooltip>
          {calcMode === 'manual' && selection?.modelId && (
            <Tooltip title={t('app.calculateAll')}>
              <Button
                size="small"
                variant="outlined"
                disabled={calcRunning}
                startIcon={calcRunning ? <CircularProgress size={12} /> : <CalculateOutlined fontSize="small" />}
                onClick={async () => {
                  const startedAt = Date.now()
                  setCalcRunning(true)
                  setCalcProgress({ done: 0, total: 1, startedAt })
                  await api.calculateModelStream(selection.modelId, (data) => {
                    if (data.phase === 'start') {
                      setCalcProgress({
                        done: 0, total: data.total_sheets || 1,
                        totalCells: data.total_cells ?? undefined,
                        computed: 0, startedAt,
                      })
                    } else if (data.phase === 'sheet_done') {
                      setCalcProgress({
                        done: data.done || 0, total: data.total_sheets || 1,
                        sheet: data.sheet,
                        totalCells: data.total_cells ?? undefined,
                        computed: data.computed ?? undefined,
                        startedAt,
                      })
                    } else if (data.phase === 'done') {
                      setCalcProgress(null); setCalcRunning(false)
                    }
                  }).catch(() => { setCalcRunning(false); setCalcProgress(null) })
                }}
                sx={{ fontSize: 11, textTransform: 'none', minWidth: 0, py: 0, px: 1 }}
              >
                {calcRunning && calcProgress ? `${calcProgress.done}/${calcProgress.total}` : t('app.calculate')}
              </Button>
            </Tooltip>
          )}

          {/* Analytics order button */}
          {isSheetSelected && (
            <Tooltip title={t('app.analyticsOrder', 'Порядок аналитик')}>
              <IconButton size="small" onClick={() => {
                const info = gridRef.current?.getAnalyticsInfo()
                if (info) {
                  setReorderItems(info.order)
                  setReorderNames(info.names)
                  setReorderOpen(true)
                }
              }}>
                <FormatListNumberedOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}

          {/* Undo button + dropdown */}
          {isSheetSelected && (
            <>
              <Tooltip title={t('app.undo', 'Отменить')}>
                <span>
                  <IconButton size="small" disabled={!hasUndo} data-testid="undo-btn" onClick={async () => {
                    const modelId = selection?.modelId
                    if (!modelId) return
                    try {
                      const hist = await api.getModelHistory(modelId, 1)
                      if (hist.length > 0) {
                        const result = await api.undoChanges(modelId, hist[0].id)
                        if (result.all_cells && gridRef.current) {
                          gridRef.current.applyCellUpdates(result.all_cells)
                        } else {
                          setRefreshKey(k => k + 1)
                        }
                        api.getModelHistory(modelId, 1).then(h => setHasUndo(h.length > 0)).catch(() => {})
                      }
                    } catch { /* ignore */ }
                  }}>
                    <UndoOutlined fontSize="small" />
                  </IconButton>
                </span>
              </Tooltip>
              <Tooltip title={t('app.undoHistory', 'История изменений')}>
                <span>
                  <IconButton size="small" disabled={!hasUndo} data-testid="undo-dropdown-btn" onClick={async (e) => {
                    const modelId = selection?.modelId
                    if (!modelId) return
                    try {
                      const items = await api.getModelHistory(modelId, 20)
                      setUndoItems(items)
                      setUndoHoverIdx(null)
                      setUndoAnchor(e.currentTarget)
                    } catch { /* ignore */ }
                  }}>
                    <ArrowDropDownOutlined fontSize="small" />
                  </IconButton>
                </span>
              </Tooltip>
              <Popover
                open={!!undoAnchor}
                anchorEl={undoAnchor}
                onClose={() => setUndoAnchor(null)}
                anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }}
                transformOrigin={{ vertical: 'top', horizontal: 'right' }}
              >
                <Box sx={{ maxHeight: 320, overflowY: 'auto', minWidth: 300 }}>
                  {undoItems.length === 0 && (
                    <Typography sx={{ p: 2, fontSize: 12, color: '#999' }}>{t('app.noHistory', 'Нет истории')}</Typography>
                  )}
                  {undoItems.map((item, i) => (
                    <Box
                      key={item.id}
                      sx={{
                        px: 1.5, py: 0.7, cursor: 'pointer', fontSize: 12, borderBottom: '1px solid #f0f0f0',
                        bgcolor: undoHoverIdx !== null && i <= undoHoverIdx ? '#e3f2fd' : 'transparent',
                        '&:hover': { bgcolor: '#e3f2fd' },
                      }}
                      onMouseEnter={() => setUndoHoverIdx(i)}
                      onMouseLeave={() => setUndoHoverIdx(null)}
                      onClick={async () => {
                        setUndoAnchor(null)
                        const modelId = selection?.modelId
                        if (!modelId) return
                        try {
                          const result = await api.undoChanges(modelId, item.id)
                          if (result.all_cells && gridRef.current) {
                            gridRef.current.applyCellUpdates(result.all_cells)
                          } else {
                            setRefreshKey(k => k + 1)
                          }
                          api.getModelHistory(modelId, 1).then(h => setHasUndo(h.length > 0)).catch(() => {})
                        } catch { /* ignore */ }
                      }}
                    >
                      <Box sx={{ display: 'flex', justifyContent: 'space-between', gap: 1 }}>
                        <Typography sx={{ fontSize: 12, fontWeight: 500, flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {item.description || item.sheet_name || '?'}
                        </Typography>
                        <Typography sx={{ fontSize: 11, color: '#999', whiteSpace: 'nowrap' }}>
                          {item.created_at?.slice(11, 16) || ''}
                        </Typography>
                      </Box>
                      <Typography sx={{ fontSize: 11, color: '#666' }}>
                        {item.old_value ?? '∅'} → {item.new_value ?? '∅'}
                        {item.username ? ` · ${item.username}` : ''}
                      </Typography>
                    </Box>
                  ))}
                </Box>
              </Popover>
            </>
          )}

          <div style={{ flex: 1 }} />

          <LanguageSwitcher />

          {isAdmin && (
            <Tooltip title={t('app.deleteAllModels')}>
              <IconButton size="small" onClick={() => setConfirmDeleteAll(true)} sx={{ color: '#d32f2f' }}>
                <DeleteSweepOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}
          {isAdmin && (
            <Tooltip title={t('app.users')}>
              <IconButton size="small" onClick={() => setShowUsers(true)}>
                <PeopleOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}


          <Tooltip title={chatOpen ? t('app.hideChat') : t('app.showChat')}>
            <IconButton
              size="small"
              onClick={() => setChatOpen(v => !v)}
              data-testid="chat-toggle"
              sx={{ color: chatOpen ? '#1976d2' : undefined }}
            >
              <SmartToyOutlined fontSize="small" />
            </IconButton>
          </Tooltip>

          {authUser && (
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, ml: 1 }}>
              <Typography sx={{ fontSize: 13, color: '#555' }}>{authUser.username}</Typography>
              <Tooltip title={t('app.logout')}>
                <IconButton size="small" onClick={onLogout}>
                  <LogoutOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
            </Box>
          )}
        </div>

        <div className="app-body">
          {leftOpen && <>
            <div style={{ width: leftWidth, minWidth: 180, flexShrink: 0 }}>
              <LeftPanel
                selection={selection} onSelect={handleSelect}
                refreshKey={refreshKey} expandAfterCreate={expandAfterCreate} onCreated={onCreated}
                sheetsOnly={isDataMode} currentUserId={isDataMode ? currentUserId : undefined}
                isAdmin={isAdmin} onRefresh={onRefresh}
              />
            </div>
            <Splitter onResize={d => setLeftWidth(w => Math.max(180, w + d))} />
          </>}

          {/* Center area: chart, settings, or pivot grid */}
          <div style={{ flex: 1, display: 'flex', minWidth: 0 }}>
            {presentation ? (
              <PresentationPanel html={presentation.html} title={presentation.title} onClose={() => setPresentation(null)} />
            ) : chartConfig ? (
              <ChartPanel config={chartConfig} onClose={() => setChartConfig(null)} />
            ) : mode === 'settings' ? (
              <CenterPanel selection={selection} onRefresh={onRefresh} />
            ) : isSheetSelected ? (
              <PivotGridAG
                ref={gridRef}
                key={`ag-${selection.id}-${refreshKey}-${mode}`}
                sheetId={selection.id} modelId={selection.modelId}
                currentUserId={currentUserId}
                calcProgress={calcProgress}
                mode={mode === 'formulas' ? 'formulas' : 'data'}
                onUndoStateChange={setHasUndo}
              />
            ) : (
              <div className="panel-center" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999' }}>
                {t('app.selectSheet')}
              </div>
            )}
          </div>

          {/* Right-docked AI chat — pushes other content via flex, not overlay */}
          {chatOpen && (
            <Splitter onResize={d => setChatWidth(w => Math.max(280, Math.min(900, w - d)))} />
          )}
          <ChatPanel
            open={chatOpen}
            width={chatWidth}
            onClose={() => setChatOpen(false)}
            context={{
              current_model_id: selection?.modelId ?? null,
              current_sheet_id: isSheetSelected ? selection!.id : null,
              user_id: currentUserId || null,
            }}
            onOpenSheet={(modelId, sheetId) => {
              setSelection({ type: 'sheet', id: sheetId, modelId })
              if (mode === 'settings') setMode('data')
            }}
            onSwitchMode={m => setMode(m)}
            onImportExcel={file => setChatImportFile(file)}
            onRefreshData={onRefresh}
            onShowChart={(cfg: any) => setChartConfig({
              title: cfg.title || '',
              chart_type: cfg.chart_type || 'bar',
              data: cfg.data || [],
              series: cfg.series || [],
              category_field: cfg.category_field || 'category',
            })}
            onShowPresentation={(p: any) => { console.log('[PEBBLE] onShowPresentation called', p?.type, 'html length:', p?.html?.length); setPresentation({ html: p.html, title: p.title || t('app.presentation') }) }}
          />
        </div>

        <UsersDialog open={showUsers} onClose={() => setShowUsers(false)} />

        {/* Analytics reorder dialog */}
        <Dialog open={reorderOpen} onClose={() => setReorderOpen(false)} maxWidth="xs" fullWidth>
          <DialogTitle>{t('app.analyticsOrder', 'Порядок аналитик')}</DialogTitle>
          <DialogContent>
            <Typography variant="caption" color="textSecondary" sx={{ mb: 1, display: 'block' }}>
              {t('app.analyticsOrderHint', 'Первая = столбцы, остальные = строки (вложенность по порядку)')}
            </Typography>
            <List dense>
              {reorderItems.map((id, i) => (
                <ListItem key={id} draggable
                  onDragStart={() => { reorderDragIdx.current = i }}
                  onDragOver={e => { e.preventDefault(); setReorderDragOver(i) }}
                  onDrop={() => {
                    const from = reorderDragIdx.current
                    if (from !== null && from !== i) {
                      const n = [...reorderItems]
                      const [m] = n.splice(from, 1)
                      n.splice(i, 0, m)
                      setReorderItems(n)
                      gridRef.current?.applyAnalyticsOrder(n)
                    }
                    reorderDragIdx.current = null; setReorderDragOver(null)
                  }}
                  onDragEnd={() => { reorderDragIdx.current = null; setReorderDragOver(null) }}
                  sx={{ cursor: 'grab', borderTop: reorderDragOver === i ? '2px solid #1976d2' : '2px solid transparent' }}>
                  <ListItemIcon sx={{ minWidth: 28 }}><DragIndicatorOutlined sx={{ fontSize: 16, color: '#bbb' }} /></ListItemIcon>
                  <ListItemText primary={`${i + 1}. ${reorderNames[id] || id}`} />
                </ListItem>
              ))}
            </List>
          </DialogContent>
        </Dialog>

        <Dialog open={confirmDeleteAll} onClose={() => setConfirmDeleteAll(false)}>
          <DialogTitle>{t('app.deleteAllTitle')}</DialogTitle>
          <DialogContent>
            <Typography>{t('app.deleteAllText')}</Typography>
          </DialogContent>
          <DialogActions>
            <Button onClick={() => setConfirmDeleteAll(false)} disabled={deletingAll}>{t('app.cancel')}</Button>
            <Button onClick={handleDeleteAllModels} color="error" variant="contained" disabled={deletingAll}>
              {deletingAll ? <CircularProgress size={20} /> : t('app.deleteAll')}
            </Button>
          </DialogActions>
        </Dialog>
        <ImportDialog open={showImport} onClose={() => setShowImport(false)} onImported={handleImported} />
        {/* Excel dropped into chat — open the import dialog pre-filled */}
        {chatImportFile && (
          <ImportDialog
            open={true}
            onClose={() => setChatImportFile(null)}
            onImported={(mid) => { handleImported(mid); setChatImportFile(null) }}
            initialFile={chatImportFile}
          />
        )}
      </div>
    </PendingProvider>
  )
}

export default function App() {
  const [auth, setAuth] = useState<{ token: string; user: { id: string; username: string; can_admin: boolean } } | null>(() => {
    const token = localStorage.getItem('pebble_token')
    const user = localStorage.getItem('pebble_user')
    if (token && user) return { token, user: JSON.parse(user) }
    return null
  })

  if (!auth) {
    return <LoginPage onLogin={(token, user) => setAuth({ token, user })} />
  }

  return <AppInner authUser={auth.user} onLogout={() => {
    localStorage.removeItem('pebble_token')
    localStorage.removeItem('pebble_user')
    setAuth(null)
  }} />
}
