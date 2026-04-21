import { useState, useRef, useEffect, useCallback } from 'react'
import { Box, IconButton, TextField, CircularProgress, Typography, Tooltip } from '@mui/material'
import SendOutlined from '@mui/icons-material/SendOutlined'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import ClearAllOutlined from '@mui/icons-material/ClearAllOutlined'
import MicOutlined from '@mui/icons-material/MicOutlined'
import MicOffOutlined from '@mui/icons-material/MicOffOutlined'
import * as api from '../../api'
import type { ChatAction } from '../../api'

export interface ChatContext {
  current_model_id?: string | null
  current_sheet_id?: string | null
  user_id?: string | null
}

export interface ChatPanelProps {
  open: boolean
  width?: number
  onClose: () => void
  context: ChatContext
  // Callbacks the agent can invoke via action side-effects
  onOpenSheet?: (modelId: string, sheetId: string) => void
  onSwitchMode?: (mode: 'settings' | 'data' | 'formulas') => void
  onImportExcel?: (file: File) => void            // user dropped an Excel file in the chat
  onRefreshData?: () => void                       // backend tool changed DB — reload
  onShowChart?: (config: any) => void              // agent built a chart — show it
}

interface UIMessage {
  role: 'user' | 'assistant'
  text: string
  // Keep raw backend content for multi-turn tool calls. User messages keep `content` = text;
  // assistant messages get the full content array so tool_use blocks round-trip correctly.
  raw: any
}

const STORAGE_KEY = 'pebble_chat_history_v1'

export default function ChatPanel({
  open, width, onClose, context,
  onOpenSheet, onSwitchMode, onImportExcel, onRefreshData, onShowChart,
}: ChatPanelProps) {
  const [messages, setMessages] = useState<UIMessage[]>(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY)
      if (raw) return JSON.parse(raw)
    } catch { /* ignore */ }
    return []
  })
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const [dragOver, setDragOver] = useState(false)
  const scrollRef = useRef<HTMLDivElement>(null)
  // Voice input state
  const [listening, setListening] = useState(false)
  const [voiceUnsupported, setVoiceUnsupported] = useState(false)
  const [micLevel, setMicLevel] = useState(0) // 0..1 audio level
  const recognitionRef = useRef<any>(null)
  const sendRef = useRef<(text: string) => void>(() => {})
  const pauseTimerRef = useRef<number | null>(null)
  const audioCtxRef = useRef<AudioContext | null>(null)
  const micStreamRef = useRef<MediaStream | null>(null)
  const animFrameRef = useRef<number>(0)
  // TTS bookkeeping: index of the last assistant message we've already spoken.
  const lastSpokenRef = useRef<number>(-1)

  useEffect(() => {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(messages.slice(-50))) } catch { /* ignore */ }
  }, [messages])

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: 'smooth' })
  }, [messages, loading])

  const applyActions = useCallback((actions: ChatAction[]) => {
    let needReload = false
    for (const a of actions) {
      if (a.type === 'open_sheet' && onOpenSheet) onOpenSheet(a.model_id, a.sheet_id)
      else if (a.type === 'switch_mode' && onSwitchMode) onSwitchMode(a.mode)
      else if (a.type === 'show_chart' && onShowChart) onShowChart(a)
      else if (a.type === 'reload_sheet' || a.type === 'reload_model') needReload = true
    }
    if (needReload && onRefreshData) onRefreshData()
  }, [onOpenSheet, onSwitchMode, onRefreshData, onShowChart])

  const send = useCallback(async (text: string) => {
    const userMsg: UIMessage = { role: 'user', text, raw: { role: 'user', content: text } }
    const nextMsgs = [...messages, userMsg]
    setMessages(nextMsgs)
    setInput('')
    setLoading(true)
    try {
      const payload = nextMsgs.map(m => ({ role: m.role, content: m.raw.content }))
      const resp = await api.chatMessage(payload, context)
      const assistantMsg: UIMessage = {
        role: 'assistant',
        text: resp.message || '(пусто)',
        raw: { role: 'assistant', content: resp.message || '' },
      }
      setMessages(prev => [...prev, assistantMsg])
      if (resp.actions?.length) applyActions(resp.actions)
    } catch (e: any) {
      setMessages(prev => [...prev, {
        role: 'assistant',
        text: `⚠️ Ошибка: ${e.message || e}`,
        raw: { role: 'assistant', content: `Ошибка: ${e.message || e}` },
      }])
    } finally {
      setLoading(false)
    }
  }, [messages, context, applyActions])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setDragOver(false)
    const file = e.dataTransfer.files?.[0]
    if (!file) return
    const isExcel = /\.xlsx?$/i.test(file.name)
    if (!isExcel) {
      setMessages(prev => [...prev, {
        role: 'assistant',
        text: `⚠️ Поддерживаются только файлы .xlsx / .xls`,
        raw: { role: 'assistant', content: '' },
      }])
      return
    }
    if (onImportExcel) {
      setMessages(prev => [...prev, {
        role: 'user',
        text: `📎 ${file.name} (импорт)`,
        raw: { role: 'user', content: `Импорт файла ${file.name}` },
      }])
      onImportExcel(file)
    }
  }, [onImportExcel])

  const clear = () => {
    setMessages([])
    localStorage.removeItem(STORAGE_KEY)
    // Also reset any live TTS and the spoken-index cursor, so re-asking the
    // same thing after a clear speaks again instead of being skipped.
    try { window.speechSynthesis?.cancel() } catch { /* ignore */ }
    lastSpokenRef.current = -1
  }

  // Keep a stable reference to `send` so voice recognition callbacks
  // always hit the latest closure without reattaching handlers on every render.
  useEffect(() => { sendRef.current = send }, [send])

  // Mic level meter via Web Audio API
  const startMicMeter = useCallback(async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true })
      micStreamRef.current = stream
      const ctx = new AudioContext()
      audioCtxRef.current = ctx
      const source = ctx.createMediaStreamSource(stream)
      const analyser = ctx.createAnalyser()
      analyser.fftSize = 256
      source.connect(analyser)
      const buf = new Uint8Array(analyser.frequencyBinCount)
      const tick = () => {
        analyser.getByteFrequencyData(buf)
        let sum = 0
        for (let i = 0; i < buf.length; i++) sum += buf[i]
        setMicLevel(Math.min(1, (sum / buf.length) / 128))
        animFrameRef.current = requestAnimationFrame(tick)
      }
      tick()
    } catch { /* mic access denied — speech recognition handles this */ }
  }, [])

  const stopMicMeter = useCallback(() => {
    if (animFrameRef.current) cancelAnimationFrame(animFrameRef.current)
    animFrameRef.current = 0
    audioCtxRef.current?.close()
    audioCtxRef.current = null
    micStreamRef.current?.getTracks().forEach(t => t.stop())
    micStreamRef.current = null
    setMicLevel(0)
  }, [])

  // Voice input: Web Speech API. Recognition runs continuously; after a
  // ~1.2 s silence we auto-submit the accumulated transcript and keep the
  // mic open for the next utterance (so user can dictate multiple commands).
  const toggleVoice = useCallback(() => {
    const W = window as any
    const SR = W.SpeechRecognition || W.webkitSpeechRecognition
    if (!SR) { setVoiceUnsupported(true); return }
    if (listening) {
      recognitionRef.current?.stop?.()
      recognitionRef.current = null
      setListening(false)
      stopMicMeter()
      if (pauseTimerRef.current) {
        window.clearTimeout(pauseTimerRef.current)
        pauseTimerRef.current = null
      }
      return
    }
    const rec = new SR()
    rec.lang = 'ru-RU'
    rec.interimResults = true
    rec.continuous = true
    let buffer = ''
    let lastInterim = ''
    const commitNow = () => {
      // Merge any interim text still shown into the buffer so we don't lose
      // words the engine hasn't finalized yet.
      const combined = (buffer + lastInterim).trim()
      if (combined) {
        sendRef.current(combined)
        buffer = ''
        lastInterim = ''
        setInput('')
      }
    }
    const schedulePauseCommit = () => {
      if (pauseTimerRef.current) window.clearTimeout(pauseTimerRef.current)
      pauseTimerRef.current = window.setTimeout(commitNow, 1200)
    }
    rec.onresult = (e: any) => {
      let interim = ''
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const r = e.results[i]
        const t = r[0]?.transcript || ''
        if (r.isFinal) buffer += t + ' '
        else interim += t
      }
      lastInterim = interim
      setInput(buffer + interim)
      // Reschedule on EVERY result event — final or interim. If silence
      // follows for >1.2s, commit fires regardless of final/interim state.
      schedulePauseCommit()
    }
    rec.onerror = (ev: any) => {
      const err = ev?.error || 'unknown'
      console.warn('[voice] SpeechRecognition error:', err)
      if (err === 'not-allowed' || err === 'service-not-allowed') {
        // Microphone permission denied
        setVoiceUnsupported(true)
        recognitionRef.current = null
        setListening(false)
      }
      // 'no-speech' and 'aborted' are transient — onend will restart
    }
    rec.onend = () => {
      // If still listening (user didn't toggle off), restart so it keeps going.
      if (recognitionRef.current === rec) {
        try { rec.start() } catch (e) { console.warn('[voice] restart failed:', e) }
      }
    }
    recognitionRef.current = rec
    setListening(true)
    startMicMeter()
    try { rec.start() } catch (e) {
      console.error('[voice] start failed:', e)
      setListening(false)
      stopMicMeter()
    }
  }, [listening, startMicMeter, stopMicMeter])

  // Cleanup on unmount
  useEffect(() => () => {
    recognitionRef.current?.stop?.()
    if (pauseTimerRef.current) window.clearTimeout(pauseTimerRef.current)
    stopMicMeter()
  }, [stopMicMeter])

  // Global shortcut: App.tsx dispatches 'pebble:toggleVoice' on double-space.
  useEffect(() => {
    const h = () => toggleVoice()
    window.addEventListener('pebble:toggleVoice', h)
    return () => window.removeEventListener('pebble:toggleVoice', h)
  }, [toggleVoice])

  // TTS voiceover of assistant replies removed per user request.

  if (!open) return null

  return (
    <Box
      data-testid="chat-panel"
      sx={{
        width: width ?? 400, flexShrink: 0, height: '100%',
        display: 'flex', flexDirection: 'column',
        borderLeft: '1px solid #e0e0e0', background: '#fafafa',
      }}
      onDragOver={e => { e.preventDefault(); setDragOver(true) }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
    >
      {/* Header */}
      <Box sx={{
        display: 'flex', alignItems: 'center', px: 1.5, py: 1,
        borderBottom: '1px solid #e0e0e0', background: '#fff',
      }}>
        <Typography sx={{ flex: 1, fontSize: 14, fontWeight: 600 }}>AI-помощник</Typography>
        <Tooltip title="Очистить историю">
          <IconButton size="small" onClick={clear}><ClearAllOutlined fontSize="small" /></IconButton>
        </Tooltip>
        <Tooltip title="Закрыть">
          <IconButton size="small" onClick={onClose}><CloseOutlined fontSize="small" /></IconButton>
        </Tooltip>
      </Box>

      {/* Messages */}
      <Box ref={scrollRef} sx={{ flex: 1, overflowY: 'auto', p: 1.5, position: 'relative' }}>
        {messages.length === 0 && (
          <Typography sx={{ color: '#888', fontSize: 12, fontStyle: 'italic' }}>
            Задавайте вопросы, просите создать модель, открыть лист, ввести значения.
            Можно перетащить Excel-файл сюда для импорта.
          </Typography>
        )}
        {messages.map((m, i) => (
          <Box key={i} sx={{
            mb: 1, display: 'flex',
            justifyContent: m.role === 'user' ? 'flex-end' : 'flex-start',
          }}>
            <Box sx={{
              maxWidth: '90%', px: 1.25, py: 0.75, borderRadius: 1.5,
              fontSize: 13, whiteSpace: 'pre-wrap', wordBreak: 'break-word',
              background: m.role === 'user' ? '#1976d2' : '#fff',
              color: m.role === 'user' ? '#fff' : '#333',
              border: m.role === 'user' ? 'none' : '1px solid #e0e0e0',
            }}>
              {m.text}
            </Box>
          </Box>
        ))}
        {loading && (
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, color: '#888', fontSize: 12 }}>
            <CircularProgress size={14} /> думаю…
          </Box>
        )}
        {dragOver && (
          <Box sx={{
            position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: 'rgba(25,118,210,0.08)', border: '2px dashed #1976d2',
            color: '#1976d2', fontSize: 14, fontWeight: 600, pointerEvents: 'none',
          }}>
            Отпустите Excel-файл для импорта
          </Box>
        )}
      </Box>

      {/* Input */}
      <Box sx={{ display: 'flex', flexDirection: 'column', borderTop: '1px solid #e0e0e0', background: '#fff' }}>
        <Box sx={{ p: 1, display: 'flex', gap: 0.5 }}>
          <TextField
            value={input}
            onChange={e => setInput(e.target.value)}
            placeholder="Задайте вопрос или команду…"
            multiline
            minRows={1}
            maxRows={4}
            size="small"
            fullWidth
            disabled={loading}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault()
                const v = input.trim()
                if (v) send(v)
              }
            }}
            inputProps={{ 'data-testid': 'chat-input' }}
          />
          <IconButton
            onClick={() => { const v = input.trim(); if (v) send(v) }}
            disabled={loading || !input.trim()}
            data-testid="chat-send"
            color="primary"
          >
            <SendOutlined />
          </IconButton>
        </Box>
        {/* Mode bar: voice toggle */}
        <Box sx={{
          px: 1, pb: 0.75, display: 'flex', alignItems: 'center', gap: 0.75,
          fontSize: 11, color: '#666',
        }}>
          <Tooltip title={
            voiceUnsupported
              ? 'Голосовой ввод не поддерживается в этом браузере'
              : listening ? 'Выключить голосовой ввод (⎵⎵)' : 'Голосовой ввод (⎵⎵ — двойной пробел)'
          }>
            <span>
              <IconButton
                size="small"
                onClick={toggleVoice}
                disabled={voiceUnsupported}
                data-testid="chat-voice-toggle"
                sx={{
                  color: listening ? '#d32f2f' : '#666',
                  animation: listening ? 'pulse 1.2s infinite' : 'none',
                  '@keyframes pulse': {
                    '0%,100%': { opacity: 1 },
                    '50%': { opacity: 0.45 },
                  },
                }}
              >
                {listening ? <MicOutlined fontSize="small" /> : <MicOffOutlined fontSize="small" />}
              </IconButton>
            </span>
          </Tooltip>
          <span>
            {voiceUnsupported ? 'голос недоступен'
              : listening ? 'слушаю… (пауза → отправка)'
              : 'голос выключен'}
          </span>
          {listening && (
            <Box sx={{
              flex: 1, height: 6, borderRadius: 3, background: '#e0e0e0',
              overflow: 'hidden', ml: 0.5,
            }}>
              <Box sx={{
                height: '100%', borderRadius: 3,
                width: `${Math.round(micLevel * 100)}%`,
                background: micLevel > 0.5 ? '#4caf50' : micLevel > 0.15 ? '#ff9800' : '#bdbdbd',
                transition: 'width 60ms linear, background 120ms',
              }} />
            </Box>
          )}
        </Box>
      </Box>
    </Box>
  )
}
