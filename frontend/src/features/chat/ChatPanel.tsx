import { useState, useRef, useEffect, useCallback } from 'react'
import { Box, IconButton, TextField, CircularProgress, Typography, Tooltip } from '@mui/material'
import SendOutlined from '@mui/icons-material/SendOutlined'
import StopOutlined from '@mui/icons-material/StopOutlined'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import ClearAllOutlined from '@mui/icons-material/ClearAllOutlined'
import ContentCopyOutlined from '@mui/icons-material/ContentCopyOutlined'
import CheckOutlined from '@mui/icons-material/CheckOutlined'
import AttachFileOutlined from '@mui/icons-material/AttachFileOutlined'
import ImageOutlined from '@mui/icons-material/ImageOutlined'
import MicOutlined from '@mui/icons-material/MicOutlined'
import MicOffOutlined from '@mui/icons-material/MicOffOutlined'
import { Chip } from '@mui/material'
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
  onShowPresentation?: (config: any) => void       // agent built a presentation — show it
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
  onOpenSheet, onSwitchMode, onImportExcel, onRefreshData, onShowChart, onShowPresentation,
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
  const [thinkingSteps, setThinkingSteps] = useState<string[]>([])
  const [typingText, setTypingText] = useState<string | null>(null) // typewriter buffer
  const [dragOver, setDragOver] = useState(false)
  const [copiedIdx, setCopiedIdx] = useState<number | null>(null) // flash check icon
  const [attachments, setAttachments] = useState<File[]>([])
  const scrollRef = useRef<HTMLDivElement>(null)
  const chatInputRef = useRef<HTMLInputElement>(null)
  const typingRef = useRef<number>(0) // animation frame id
  const abortRef = useRef<AbortController | null>(null)
  // Prompt history index for ArrowUp/Down navigation
  const historyPosRef = useRef(-1) // -1 = current input, 0..N = index into user messages (reversed)
  const savedInputRef = useRef('')
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

  // Auto-focus chat input when panel opens
  useEffect(() => {
    if (open) setTimeout(() => chatInputRef.current?.focus(), 50)
  }, [open])

  const triggerFilePicker = useCallback(() => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.xlsx,.xls'
    input.onchange = () => {
      const file = input.files?.[0]
      if (file && onImportExcel) {
        setMessages(prev => [...prev, {
          role: 'user',
          text: `📎 ${file.name} (импорт)`,
          raw: { role: 'user', content: `Импорт файла ${file.name}` },
        }])
        onImportExcel(file)
      }
    }
    input.click()
  }, [onImportExcel])

  const applyActions = useCallback((actions: ChatAction[]) => {
    let needReload = false
    for (const a of actions) {
      if (a.type === 'open_sheet' && onOpenSheet) onOpenSheet(a.model_id, a.sheet_id)
      else if (a.type === 'switch_mode' && onSwitchMode) onSwitchMode(a.mode)
      else if (a.type === 'show_chart' && onShowChart) onShowChart(a)
      else if (a.type === 'show_presentation' && onShowPresentation) onShowPresentation(a)
      else if (a.type === 'pick_excel_file') triggerFilePicker()
      else if (a.type === 'reload_sheet' || a.type === 'reload_model') needReload = true
    }
    if (needReload && onRefreshData) onRefreshData()
  }, [onOpenSheet, onSwitchMode, onRefreshData, onShowChart, onShowPresentation, triggerFilePicker])

  const fileToBase64 = (file: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve((reader.result as string).split(',')[1])
      reader.onerror = reject
      reader.readAsDataURL(file)
    })

  const addAttachments = useCallback((files: FileList | File[]) => {
    const arr = Array.from(files)
    for (const f of arr) {
      if (/\.xlsx?$/i.test(f.name) && onImportExcel) {
        // Excel files → import flow
        setMessages(prev => [...prev, {
          role: 'user',
          text: `📎 ${f.name} (импорт)`,
          raw: { role: 'user', content: `Импорт файла ${f.name}` },
        }])
        onImportExcel(f)
        return
      }
    }
    // Non-Excel files → attach as chips
    setAttachments(prev => [...prev, ...arr])
  }, [onImportExcel])

  // Typewriter: gradually reveal full text
  const typewrite = useCallback((fullText: string, rawContent: any) => {
    let pos = 0
    const CHARS_PER_TICK = 2
    const TICK_MS = 20
    const tick = () => {
      pos = Math.min(pos + CHARS_PER_TICK, fullText.length)
      setTypingText(fullText.slice(0, pos))
      if (pos < fullText.length) {
        typingRef.current = window.setTimeout(tick, TICK_MS)
      } else {
        // Done typing — commit as real message
        setTypingText(null)
        setMessages(prev => [...prev, {
          role: 'assistant',
          text: fullText,
          raw: { role: 'assistant', content: rawContent },
        }])
      }
    }
    tick()
  }, [])

  // Cleanup typewriter on unmount
  useEffect(() => () => { if (typingRef.current) clearTimeout(typingRef.current) }, [])

  const send = useCallback(async (text: string) => {
    // Local commands — no need to hit the server
    const lower = text.trim().toLowerCase()
    if (/очист|сброс|clear/.test(lower) && /истори|чат|chat|history/.test(lower)) {
      setMessages([])
      setInput('')
      localStorage.removeItem(STORAGE_KEY)
      return
    }
    historyPosRef.current = -1
    savedInputRef.current = ''

    // Build content: text + image attachments
    const currentAttachments = [...attachments]
    const imageFiles = currentAttachments.filter(f => f.type.startsWith('image/'))
    let rawContent: any = text
    if (imageFiles.length > 0) {
      const contentParts: any[] = []
      for (const img of imageFiles) {
        const b64 = await fileToBase64(img)
        const mediaType = img.type || 'image/png'
        contentParts.push({ type: 'image', source: { type: 'base64', media_type: mediaType, data: b64 } })
      }
      contentParts.push({ type: 'text', text })
      rawContent = contentParts
    }
    const chipNames = currentAttachments.map(f => `📎 ${f.name}`).join(', ')
    const displayText = chipNames ? `${chipNames}\n${text}` : text

    const userMsg: UIMessage = { role: 'user', text: displayText, raw: { role: 'user', content: rawContent } }
    const nextMsgs = [...messages, userMsg]
    setMessages(nextMsgs)
    setInput('')
    setAttachments([])
    setLoading(true)
    setThinkingSteps([])
    const controller = new AbortController()
    abortRef.current = controller
    try {
      const payload = nextMsgs.map(m => ({ role: m.role, content: m.raw.content }))
      await api.chatMessageStream(payload, context, (event) => {
        if (event.type === 'thinking') {
          setThinkingSteps(prev => [...prev, event.text])
        } else if (event.type === 'done') {
          setLoading(false)
          setThinkingSteps([])
          const finalText = event.text || '(пусто)'
          typewrite(finalText, finalText)
          if (event.actions?.length) applyActions(event.actions)
        } else if (event.type === 'error') {
          setLoading(false)
          setThinkingSteps([])
          setMessages(prev => [...prev, {
            role: 'assistant',
            text: `⚠️ Ошибка: ${event.text}`,
            raw: { role: 'assistant', content: `Ошибка: ${event.text}` },
          }])
        }
      }, controller.signal)
    } catch (e: any) {
      if (e.name === 'AbortError' || controller.signal.aborted) {
        // User clicked stop — just clear loading state
        setLoading(false)
        setThinkingSteps([])
        return
      }
      setMessages(prev => [...prev, {
        role: 'assistant',
        text: `⚠️ Ошибка: ${e.message || e}`,
        raw: { role: 'assistant', content: `Ошибка: ${e.message || e}` },
      }])
      setLoading(false)
      setThinkingSteps([])
    }
  }, [messages, context, applyActions, typewrite, attachments])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setDragOver(false)
    const files = e.dataTransfer.files
    if (!files?.length) return
    addAttachments(files)
  }, [addAttachments])

  const handlePaste = useCallback((e: React.ClipboardEvent) => {
    const items = e.clipboardData?.items
    if (!items) return
    const files: File[] = []
    for (let i = 0; i < items.length; i++) {
      const item = items[i]
      if (item.kind === 'file') {
        const f = item.getAsFile()
        if (f) files.push(f)
      }
    }
    if (files.length > 0) {
      e.preventDefault()
      addAttachments(files)
    }
  }, [addAttachments])

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
    } catch (e) { console.warn('[mic-meter] failed:', e) }
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
        // Permission denied — stop listening but don't permanently disable
        recognitionRef.current = null
        setListening(false)
        stopMicMeter()
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
    // Start mic meter FIRST (getUserMedia), then speech recognition.
    // This ensures AudioContext gets mic access before SpeechRecognition grabs it.
    startMicMeter().then(() => {
      try {
        rec.start()
      } catch (e) {
        console.error('[voice] start failed:', e)
        setListening(false)
        stopMicMeter()
      }
    }).catch(() => {
      // Mic meter failed — still try speech recognition
      try { rec.start() } catch (e) {
        console.error('[voice] start failed:', e)
        setListening(false)
      }
    })
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
            '&:hover .copy-btn': { opacity: 1 },
          }}>
            <Box sx={{
              maxWidth: '90%', position: 'relative',
              px: 1.25, py: 0.75, borderRadius: 1.5,
              fontSize: 13, whiteSpace: 'pre-wrap', wordBreak: 'break-word',
              background: m.role === 'user' ? '#1976d2' : '#fff',
              color: m.role === 'user' ? '#fff' : '#333',
              border: m.role === 'user' ? 'none' : '1px solid #e0e0e0',
            }}>
              {m.text}
              <IconButton
                className="copy-btn"
                size="small"
                sx={{
                  position: 'absolute', top: 2,
                  ...(m.role === 'user' ? { left: -28 } : { right: -28 }),
                  opacity: 0, transition: 'opacity 0.15s',
                  color: m.role === 'user' ? 'rgba(255,255,255,0.6)' : '#999',
                  '&:hover': { color: m.role === 'user' ? '#fff' : '#555' },
                }}
                onClick={() => {
                  navigator.clipboard.writeText(m.text)
                  setCopiedIdx(i)
                  setTimeout(() => setCopiedIdx(null), 1500)
                }}
              >
                {copiedIdx === i
                  ? <CheckOutlined sx={{ fontSize: 14 }} />
                  : <ContentCopyOutlined sx={{ fontSize: 14 }} />}
              </IconButton>
            </Box>
          </Box>
        ))}
        {/* Typewriter: assistant message being revealed gradually */}
        {typingText !== null && (
          <Box sx={{ mb: 1, display: 'flex', justifyContent: 'flex-start' }}>
            <Box sx={{
              maxWidth: '90%', px: 1.25, py: 0.75, borderRadius: 1.5,
              fontSize: 13, whiteSpace: 'pre-wrap', wordBreak: 'break-word',
              background: '#fff', color: '#333', border: '1px solid #e0e0e0',
            }}>
              {typingText}<Box component="span" sx={{ animation: 'blink 0.7s infinite', '@keyframes blink': { '0%,100%': { opacity: 1 }, '50%': { opacity: 0 } } }}>▊</Box>
            </Box>
          </Box>
        )}
        {loading && (
          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.5, color: '#888', fontSize: 12 }}>
            {thinkingSteps.length > 0 ? (
              thinkingSteps.map((step, i) => (
                <Box key={i} sx={{
                  display: 'flex', alignItems: 'center', gap: 0.5,
                  opacity: i === thinkingSteps.length - 1 ? 1 : 0.5,
                }}>
                  {i === thinkingSteps.length - 1 && <CircularProgress size={12} />}
                  {step}
                </Box>
              ))
            ) : (
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                <CircularProgress size={14} /> думаю…
              </Box>
            )}
          </Box>
        )}
        {dragOver && (
          <Box sx={{
            position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: 'rgba(25,118,210,0.08)', border: '2px dashed #1976d2',
            color: '#1976d2', fontSize: 14, fontWeight: 600, pointerEvents: 'none',
          }}>
            Отпустите файл для прикрепления
          </Box>
        )}
      </Box>

      {/* Input */}
      <Box sx={{ display: 'flex', flexDirection: 'column', borderTop: '1px solid #e0e0e0', background: '#fff' }}>
        {attachments.length > 0 && (
          <Box sx={{ px: 1, pt: 0.75, display: 'flex', flexWrap: 'wrap', gap: 0.5 }}>
            {attachments.map((f, i) => (
              <Chip
                key={i}
                label={f.name}
                size="small"
                icon={f.type.startsWith('image/') ? <ImageOutlined sx={{ fontSize: 16 }} /> : <AttachFileOutlined sx={{ fontSize: 16 }} />}
                onDelete={() => setAttachments(prev => prev.filter((_, j) => j !== i))}
                sx={{ fontSize: 11, maxWidth: 180 }}
              />
            ))}
          </Box>
        )}
        <Box sx={{ p: 1, display: 'flex', gap: 0.5 }}>
          <TextField
            value={input}
            onChange={e => setInput(e.target.value)}
            onPaste={handlePaste}
            placeholder="Задайте вопрос или команду…"
            multiline
            minRows={1}
            maxRows={4}
            size="small"
            fullWidth
            disabled={loading}
            inputRef={chatInputRef}
            onKeyDown={e => {
              if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault()
                const v = input.trim()
                if (v) send(v)
              } else if (e.key === 'ArrowUp' && !e.shiftKey) {
                // Derive history from all user messages (oldest first)
                const hist = messages.filter(m => m.role === 'user').map(m => m.text)
                if (hist.length === 0) return
                const el = e.target as HTMLTextAreaElement
                if (el.selectionStart === 0 && el.selectionEnd === 0) {
                  e.preventDefault()
                  if (historyPosRef.current < 0) {
                    savedInputRef.current = input
                    historyPosRef.current = hist.length - 1
                  } else if (historyPosRef.current > 0) {
                    historyPosRef.current--
                  }
                  setInput(hist[historyPosRef.current])
                }
              } else if (e.key === 'ArrowDown' && !e.shiftKey) {
                if (historyPosRef.current < 0) return
                const hist = messages.filter(m => m.role === 'user').map(m => m.text)
                const el = e.target as HTMLTextAreaElement
                const atEnd = el.selectionStart === el.value.length
                if (atEnd) {
                  e.preventDefault()
                  if (historyPosRef.current < hist.length - 1) {
                    historyPosRef.current++
                    setInput(hist[historyPosRef.current])
                  } else {
                    historyPosRef.current = -1
                    setInput(savedInputRef.current)
                  }
                }
              }
            }}
            inputProps={{ 'data-testid': 'chat-input' }}
          />
          {loading ? (
            <IconButton
              onClick={() => abortRef.current?.abort()}
              data-testid="chat-stop"
              color="error"
            >
              <StopOutlined />
            </IconButton>
          ) : (
            <IconButton
              onClick={() => { const v = input.trim(); if (v) send(v) }}
              disabled={!input.trim()}
              data-testid="chat-send"
              color="primary"
            >
              <SendOutlined />
            </IconButton>
          )}
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
