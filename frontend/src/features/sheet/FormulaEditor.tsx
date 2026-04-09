import { useState, useEffect, useRef, useCallback } from 'react'
import {
  Dialog, DialogTitle, DialogContent, DialogActions, Button,
  Box, Typography, TextField, Chip, Divider,
} from '@mui/material'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import ChevronRightOutlined from '@mui/icons-material/ChevronRightOutlined'
import DescriptionOutlined from '@mui/icons-material/DescriptionOutlined'
import * as api from '../../api'
import type { Sheet, Analytic, AnalyticRecord } from '../../types'

interface RecordNode {
  record: AnalyticRecord; data: Record<string, any>; children: RecordNode[]
}

function buildTree(records: AnalyticRecord[]): RecordNode[] {
  const byParent: Record<string, AnalyticRecord[]> = { root: [] }
  for (const r of records) (byParent[r.parent_id || 'root'] ||= []).push(r)
  const build = (pid: string | null): RecordNode[] =>
    (byParent[pid || 'root'] || []).map(r => ({
      record: r,
      data: typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json,
      children: build(r.id),
    }))
  return build(null)
}

interface SheetTree {
  sheet: Sheet
  analytics: { analytic: Analytic; records: RecordNode[] }[]
}

// ─── Expression templates ───
const TEMPLATES = [
  {
    label: 'если...то...иначе',
    code: 'если [условие] то\n  [значение]\nиначе\n  [значение]\nконец_если',
  },
  {
    label: 'если...то...иначе_если...иначе',
    code: 'если [условие] то\n  [значение]\nиначе_если [условие] то\n  [значение]\nиначе\n  [значение]\nконец_если',
  },
]

interface Props {
  open: boolean
  formula: string
  onSave: (formula: string) => void
  onClose: () => void
  modelId: string
}

export default function FormulaEditor({ open, formula, onSave, onClose, modelId }: Props) {
  const [text, setText] = useState(formula)
  const [sheets, setSheets] = useState<SheetTree[]>([])
  const [expanded, setExpanded] = useState<Set<string>>(new Set())
  const textRef = useRef<HTMLTextAreaElement>(null)

  useEffect(() => { setText(formula) }, [formula])

  useEffect(() => {
    if (!open || !modelId) return
    ;(async () => {
      const tree = await api.getModelTree(modelId)
      const sheetList: Sheet[] = tree.sheets || []
      const analyticsList: Analytic[] = tree.analytics || []

      const result: SheetTree[] = []
      for (const s of sheetList) {
        const sa = await api.listSheetAnalytics(s.id)
        const items: { analytic: Analytic; records: RecordNode[] }[] = []
        for (const binding of sa) {
          const a = analyticsList.find(x => x.id === binding.analytic_id)
          if (!a) continue
          const recs = await api.listRecords(a.id)
          items.push({ analytic: a, records: buildTree(recs) })
        }
        result.push({ sheet: s, analytics: items })
      }
      setSheets(result)
    })()
  }, [open, modelId])

  const toggle = (key: string) => {
    setExpanded(prev => {
      const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n
    })
  }

  const insertRef = (sheetName: string, analyticName: string, recordName: string) => {
    const ref = `[${sheetName}].[${analyticName}].[${recordName}]`
    const el = textRef.current
    if (el) {
      const start = el.selectionStart
      const end = el.selectionEnd
      const before = text.slice(0, start)
      const after = text.slice(end)
      const next = before + ref + after
      setText(next)
      setTimeout(() => {
        el.focus()
        el.selectionStart = el.selectionEnd = start + ref.length
      }, 0)
    } else {
      setText(prev => prev + ref)
    }
  }

  const insertTemplate = (code: string) => {
    const el = textRef.current
    if (el) {
      const start = el.selectionStart
      const before = text.slice(0, start)
      const after = text.slice(el.selectionEnd)
      setText(before + code + after)
      setTimeout(() => { el.focus() }, 0)
    } else {
      setText(prev => prev + '\n' + code)
    }
  }

  const renderRecordTree = (nodes: RecordNode[], sheetName: string, analyticName: string, level: number): React.ReactNode => {
    return nodes.map(n => {
      const key = `rec:${n.record.id}`
      const name = n.data.name || n.record.id.slice(0, 8)
      const hasChildren = n.children.length > 0
      const isExp = expanded.has(key)
      return (
        <Box key={n.record.id}>
          <Box
            sx={{
              display: 'flex', alignItems: 'center', pl: 2 + level * 2, py: 0.25,
              cursor: 'pointer', fontSize: 12, '&:hover': { bgcolor: '#e3f2fd' },
            }}
          >
            {hasChildren ? (
              <Box onClick={() => toggle(key)} sx={{ display: 'flex', mr: 0.5 }}>
                {isExp ? <ExpandMoreOutlined sx={{ fontSize: 14 }} /> : <ChevronRightOutlined sx={{ fontSize: 14 }} />}
              </Box>
            ) : <Box sx={{ width: 18 }} />}
            <Typography
              variant="body2" sx={{ fontSize: 12 }}
              onClick={() => insertRef(sheetName, analyticName, name)}
            >
              {name}
            </Typography>
          </Box>
          {hasChildren && isExp && renderRecordTree(n.children, sheetName, analyticName, level + 1)}
        </Box>
      )
    })
  }

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle sx={{ py: 1 }}>Редактор формул</DialogTitle>
      <DialogContent sx={{ display: 'flex', gap: 2, height: 450, p: 2 }}>
        {/* Left: sheets & indicators tree */}
        <Box sx={{ width: 240, flexShrink: 0, overflow: 'auto', border: '1px solid #e0e0e0', borderRadius: 1 }}>
          <Typography variant="caption" sx={{ px: 1, pt: 0.5, display: 'block', color: '#999' }}>
            Листы и показатели
          </Typography>
          {sheets.map(st => {
            const sKey = `sheet:${st.sheet.id}`
            const sExp = expanded.has(sKey)
            return (
              <Box key={st.sheet.id}>
                <Box onClick={() => toggle(sKey)}
                  sx={{ display: 'flex', alignItems: 'center', px: 1, py: 0.5, cursor: 'pointer', '&:hover': { bgcolor: '#f5f5f5' } }}>
                  {sExp ? <ExpandMoreOutlined sx={{ fontSize: 14 }} /> : <ChevronRightOutlined sx={{ fontSize: 14 }} />}
                  <DescriptionOutlined sx={{ fontSize: 14, mx: 0.5, opacity: 0.5 }} />
                  <Typography variant="body2" sx={{ fontSize: 12 }}>{st.sheet.name}</Typography>
                </Box>
                {sExp && st.analytics.map(({ analytic, records }) => {
                  const aKey = `analytic:${analytic.id}`
                  const aExp = expanded.has(aKey)
                  return (
                    <Box key={analytic.id}>
                      <Box onClick={() => toggle(aKey)}
                        sx={{ display: 'flex', alignItems: 'center', pl: 3, py: 0.25, cursor: 'pointer', '&:hover': { bgcolor: '#f5f5f5' } }}>
                        {aExp ? <ExpandMoreOutlined sx={{ fontSize: 14 }} /> : <ChevronRightOutlined sx={{ fontSize: 14 }} />}
                        <Typography variant="body2" sx={{ fontSize: 12, ml: 0.5 }}>{analytic.name}</Typography>
                      </Box>
                      {aExp && renderRecordTree(records, st.sheet.name, analytic.name, 0)}
                    </Box>
                  )
                })}
              </Box>
            )
          })}
        </Box>

        {/* Center: formula editor */}
        <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
          <Typography variant="caption" sx={{ color: '#999', mb: 0.5 }}>
            Формула (клик по показателю слева вставит ссылку)
          </Typography>
          <textarea
            ref={textRef}
            value={text}
            onChange={e => setText(e.target.value)}
            style={{
              flex: 1, fontFamily: 'monospace', fontSize: 13, padding: 8,
              border: '1px solid #e0e0e0', borderRadius: 4, outline: 'none',
              resize: 'none', lineHeight: 1.6,
            }}
            placeholder={'Пример:\nесли [Бюджет].[Показатели].[Выручка] > 0 то\n  [Бюджет].[Показатели].[Выручка] * 0.2\nиначе\n  0\nконец_если'}
          />
        </Box>

        {/* Right: expression templates */}
        <Box sx={{ width: 200, flexShrink: 0, overflow: 'auto', border: '1px solid #e0e0e0', borderRadius: 1, p: 1 }}>
          <Typography variant="caption" sx={{ color: '#999', display: 'block', mb: 1 }}>
            Выражения
          </Typography>
          {TEMPLATES.map((t, i) => (
            <Box key={i} sx={{ mb: 1 }}>
              <Chip
                label={t.label}
                size="small"
                onClick={() => insertTemplate(t.code)}
                sx={{ fontSize: 11, cursor: 'pointer' }}
              />
              <Typography variant="caption" sx={{ display: 'block', mt: 0.5, color: '#999', fontFamily: 'monospace', fontSize: 10, whiteSpace: 'pre-wrap' }}>
                {t.code}
              </Typography>
              {i < TEMPLATES.length - 1 && <Divider sx={{ mt: 1 }} />}
            </Box>
          ))}
        </Box>
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Отмена</Button>
        <Button variant="contained" onClick={() => { onSave(text); onClose() }}>Сохранить</Button>
      </DialogActions>
    </Dialog>
  )
}
