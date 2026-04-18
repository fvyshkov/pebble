import { useState, useRef, useEffect, useCallback } from 'react'
import { Box, IconButton, TextField, CircularProgress, Typography, Tooltip } from '@mui/material'
import SendOutlined from '@mui/icons-material/SendOutlined'
import CloseOutlined from '@mui/icons-material/CloseOutlined'
import ClearAllOutlined from '@mui/icons-material/ClearAllOutlined'
import * as api from '../../api'
import type { ChatAction } from '../../api'

export interface ChatContext {
  current_model_id?: string | null
  current_sheet_id?: string | null
  user_id?: string | null
}

export interface ChatPanelProps {
  open: boolean
  onClose: () => void
  context: ChatContext
  // Callbacks the agent can invoke via action side-effects
  onOpenSheet?: (modelId: string, sheetId: string) => void
  onSwitchMode?: (mode: 'settings' | 'data' | 'formulas') => void
  onImportExcel?: (file: File) => void            // user dropped an Excel file in the chat
  onRefreshData?: () => void                       // backend tool changed DB — reload
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
  open, onClose, context,
  onOpenSheet, onSwitchMode, onImportExcel, onRefreshData,
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
      else if (a.type === 'reload_sheet' || a.type === 'reload_model') needReload = true
      // pin/unpin_analytic — not wired yet; tool still records the action.
      // TODO: propagate to PivotGrid via a ViewSettings update
    }
    if (needReload && onRefreshData) onRefreshData()
  }, [onOpenSheet, onSwitchMode, onRefreshData])

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

  const clear = () => { setMessages([]); localStorage.removeItem(STORAGE_KEY) }

  if (!open) return null

  return (
    <Box
      data-testid="chat-panel"
      sx={{
        width: 400, minWidth: 320, height: '100%',
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
      <Box sx={{ p: 1, borderTop: '1px solid #e0e0e0', background: '#fff', display: 'flex', gap: 0.5 }}>
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
    </Box>
  )
}
