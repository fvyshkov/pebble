import { useState, useCallback, useEffect, useRef } from 'react'
import {
  IconButton, Tooltip, Badge, Select, MenuItem, FormControl,
  Dialog, DialogTitle, DialogContent, DialogActions, Button, TextField, CircularProgress,
  ToggleButton, ToggleButtonGroup, Typography, Box, Chip,
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
import type { TreeSelection } from './types'
import LoginPage from './features/auth/LoginPage'
import LeftPanel from './panels/LeftPanel'
import CenterPanel from './panels/CenterPanel'
import Splitter from './components/Splitter'
import UsersDialog from './components/UsersDialog'
import PivotGrid from './features/sheet/PivotGrid'
import PivotGridAG from './features/sheet/PivotGridAG'
import ChatPanel from './features/chat/ChatPanel'
import ChartPanel, { type ChartConfig } from './features/chart/ChartPanel'
import PresentationPanel from './features/presentation/PresentationPanel'
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

function ImportDialog({ open, onClose, onImported, initialFile }: {
  open: boolean; onClose: () => void; onImported: (modelId: string) => void; initialFile?: File | null
}) {
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
      })
    } catch (err) {
      setLog(prev => [...prev, `[ERR]Ошибка: ${(err as Error).message}`])
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
        {loading && !done && (
          <Box sx={{ width: '100%', mb: 1 }}>
            {progress.total > 0 ? (
              <>
                <Box sx={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#666', mb: 0.5 }}>
                  <span>Показатели: {progress.current} из {progress.total}</span>
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
                  <CircularProgress size={12} /> <span>работаю... {elapsed > 0 ? `${mins}:${secs}` : ''}</span>
                </Box>
              )
            })()}
          </Box>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={handleClose} disabled={loading} startIcon={done ? <CloseOutlined /> : undefined}>
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

function AppInner({ authUser, onLogout }: { authUser?: { id: string; username: string; can_admin: boolean }; onLogout?: () => void }) {
  const [mode, setMode] = useState<AppMode>('data')
  const [selection, setSelection] = useState<TreeSelection | null>(null)
  const [leftWidth, setLeftWidth] = useState(280)
  const [leftOpen, setLeftOpen] = useState(true)
  const [refreshKey, setRefreshKey] = useState(0)
  const [showUsers, setShowUsers] = useState(false)
  const [showImport, setShowImport] = useState(false)
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
  // AG Grid is now the only supported mode. The toggle is hidden, but the
  // state + legacy PivotGrid code path are kept for emergency fallback.
  // Intentionally no localStorage persistence: page reload always = AG Grid.
  const [useAgGrid, setUseAgGrid] = useState<boolean>(true)
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

  // ── Global shortcuts ──────────────────────────────────────────────────
  // • Double-space (two spaces within 400ms) → toggle voice input in chat.
  //   Opens chat panel if it's closed. Ignored when the user is typing in
  //   an input/textarea so normal " " still works.
  // • Cmd/Ctrl+Z → history-based undo (api.undoChanges on the latest entry).
  //   Skipped when focus is inside an editable element, so text-editor undo
  //   (including AG Grid cell editor) keeps working natively.
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
      // INPUT / TEXTAREA — use native setter so React picks up the change.
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
      // contenteditable — patch textContent if it ends with a space.
      if (el.isContentEditable) {
        const t = el.textContent || ''
        if (t.endsWith(' ')) {
          el.textContent = t.slice(0, -1)
          el.dispatchEvent(new Event('input', { bubbles: true }))
        }
      }
    }
    const onKey = async (ev: KeyboardEvent) => {
      // Double-space → toggle voice. Works everywhere in the app, including
      // inside inputs, textareas and AG Grid cell editors. In editable
      // targets we strip the just-typed trailing space so the field isn't
      // polluted by the shortcut.
      if (ev.key === ' ' && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ev.shiftKey) {
        const now = Date.now()
        if (now - lastSpaceRef.current < 400) {
          ev.preventDefault()
          lastSpaceRef.current = 0
          if (isEditable()) stripTrailingSpace()
          setChatOpen(true)
          // Defer a tick so ChatPanel has mounted if it was closed.
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
      // Cmd/Ctrl+Z → history undo
      if ((ev.metaKey || ev.ctrlKey) && !ev.shiftKey && ev.key.toLowerCase() === 'z') {
        if (isEditable()) return
        const modelId = selection?.modelId
        if (!modelId) return
        ev.preventDefault()
        try {
          const hist = await api.getModelHistory(modelId, 1)
          if (hist.length > 0) {
            await api.undoChanges(modelId, hist[0].id)
            setRefreshKey(k => k + 1)
          }
        } catch { /* ignore */ }
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [selection?.modelId])

  // No auto-calculate on first sheet load — imported values from Excel are already correct.
  // Recalculation only happens when user explicitly edits cells or clicks "Рассчитать".

  const onCreated = useCallback((info: { modelId: string; folder: 'sheets' | 'analytics'; id: string; type: 'sheet' | 'analytic' }) => {
    setExpandAfterCreate({ modelId: info.modelId, folder: info.folder, selectId: info.id, selectType: info.type })
    setSelection({ type: info.type, id: info.id, modelId: info.modelId })
    setRefreshKey(k => k + 1)
  }, [])

  const handleImported = useCallback((modelId: string) => {
    setSelection({ type: 'model', id: modelId, modelId })
    setRefreshKey(k => k + 1)
    // Auto-recalc after import — consolidation formulas need computation
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

  // When switching to data/formulas mode, if a sheet is selected — keep it.
  // When clicking a sheet in left panel — select it and switch to data mode if needed.
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
          <Tooltip title={leftOpen ? "Скрыть панель" : "Показать панель"}>
            <IconButton size="small" onClick={() => setLeftOpen(v => !v)}>
              <MenuOutlined fontSize="small" />
            </IconButton>
          </Tooltip>
          {isAdmin && (
            <Tooltip title="Импорт модели из Excel">
              <IconButton size="small" onClick={() => setShowImport(true)}>
                <FileUploadOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}
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
            <ToggleButton value="data">
              <Tooltip title="Просмотр / ввод данных"><TableChartOutlined sx={{ fontSize: 16 }} /></Tooltip>
            </ToggleButton>
            {isAdmin && (
              <ToggleButton value="formulas">
                <Tooltip title="Формулы и правила"><FunctionsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
            {isAdmin && (
              <ToggleButton value="settings">
                <Tooltip title="Настройки модели"><SettingsOutlined sx={{ fontSize: 16 }} /></Tooltip>
              </ToggleButton>
            )}
          </ToggleButtonGroup>


          {/* Calc mode toggle + calculate button */}
          <Tooltip title={calcMode === 'auto' ? 'Авто-расчёт (при каждом сохранении)' : 'Ручной расчёт (по кнопке)'}>
            <Chip
              size="small"
              label={calcMode === 'auto' ? 'авто' : 'вручную'}
              variant={calcMode === 'auto' ? 'filled' : 'outlined'}
              color={calcMode === 'auto' ? 'success' : 'default'}
              onClick={() => setCalcMode(prev => prev === 'auto' ? 'manual' : 'auto')}
              sx={{ fontSize: 11, cursor: 'pointer' }}
            />
          </Tooltip>
          {calcMode === 'manual' && selection?.modelId && (
            <Tooltip title="Рассчитать все формулы">
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
                      // NB: no longer bump refreshKey — that would remount
                      // PivotGridAG and lose the "pre-recalc" cellMap snapshot
                      // needed to detect which cells changed (for the green
                      // flash). PivotGridAG refetches + diffs internally when
                      // calcProgress transitions from truthy to null.
                      setCalcProgress(null); setCalcRunning(false)
                    }
                  }).catch(() => { setCalcRunning(false); setCalcProgress(null) })
                }}
                sx={{ fontSize: 11, textTransform: 'none', minWidth: 0, py: 0, px: 1 }}
              >
                {calcRunning && calcProgress ? `${calcProgress.done}/${calcProgress.total}` : 'Рассчитать'}
              </Button>
            </Tooltip>
          )}

          <div style={{ flex: 1 }} />

          {isAdmin && (
            <Tooltip title="Пользователи">
              <IconButton size="small" onClick={() => setShowUsers(true)}>
                <PeopleOutlined fontSize="small" />
              </IconButton>
            </Tooltip>
          )}

          {/* AG Grid toggle hidden — AG is the only active mode. Keep the
              legacy PivotGrid path compiled in case we need to fall back. */}
          {false && (
            <Tooltip title={useAgGrid ? 'Переключиться на старый grid' : 'Переключиться на AG Grid (бета)'}>
              <IconButton
                size="small"
                onClick={() => setUseAgGrid(v => !v)}
                data-testid="aggrid-toggle"
                sx={{ color: useAgGrid ? '#1976d2' : undefined, fontSize: 12 }}
              >
                <span style={{ fontWeight: 700, fontSize: 11 }}>{useAgGrid ? 'AG' : 'old'}</span>
              </IconButton>
            </Tooltip>
          )}

          <Tooltip title={chatOpen ? 'Скрыть AI-чат (⌘J)' : 'AI-помощник (⌘J)'}>
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
              <Tooltip title="Выйти">
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
                isAdmin={isAdmin}
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
              useAgGrid ? (
                <PivotGridAG
                  key={`ag-${selection.id}-${refreshKey}-${mode}`}
                  sheetId={selection.id} modelId={selection.modelId}
                  currentUserId={currentUserId}
                  calcProgress={calcProgress}
                  mode={mode === 'formulas' ? 'formulas' : 'data'}
                />
              ) : (
                <PivotGrid
                  key={`${selection.id}-${refreshKey}`}
                  sheetId={selection.id} modelId={selection.modelId}
                  currentUserId={currentUserId}
                  mode={mode === 'formulas' ? 'settings' : 'data'}
                  calcMode={calcMode}
                />
              )
            ) : (
              <div className="panel-center" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#999' }}>
                Выберите лист для просмотра данных
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
            onShowPresentation={(p: any) => { console.log('[PEBBLE] onShowPresentation called', p?.type, 'html length:', p?.html?.length); setPresentation({ html: p.html, title: p.title || 'Презентация' }) }}
          />
        </div>

        <UsersDialog open={showUsers} onClose={() => setShowUsers(false)} />
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
