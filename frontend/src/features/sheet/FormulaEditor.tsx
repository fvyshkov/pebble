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

interface AxisInfo {
  analytic: Analytic
  records: RecordNode[]
}

interface SheetTree {
  sheet: Sheet
  mainAnalytic: Analytic | null
  records: RecordNode[]
  axes: AxisInfo[]   // non-main, non-period analytics (for param insertion)
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
  currentSheetId?: string
}

export default function FormulaEditor({ open, formula, onSave, onClose, modelId, currentSheetId }: Props) {
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
        // Only the main analytic (the indicator axis). Fall back to first non-periods.
        const mainBinding = sa.find(b => (b as any).is_main === 1 || (b as any).is_main === true)
          ?? sa.find(b => {
            const a = analyticsList.find(x => x.id === b.analytic_id)
            return a && !a.is_periods
          })
        const main = mainBinding ? analyticsList.find(x => x.id === mainBinding.analytic_id) ?? null : null
        let records: RecordNode[] = []
        if (main) {
          const recs = await api.listRecords(main.id)
          records = buildTree(recs)
        }
        // Load non-main, non-period axes for param completion
        const axes: AxisInfo[] = []
        for (const b of sa) {
          const a = analyticsList.find(x => x.id === b.analytic_id)
          if (!a || a.is_periods || a.id === main?.id) continue
          const recs = await api.listRecords(a.id)
          axes.push({ analytic: a, records: buildTree(recs) })
        }
        result.push({ sheet: s, mainAnalytic: main, records, axes })
      }
      setSheets(result)
    })()
  }, [open, modelId])

  const toggle = (key: string) => {
    setExpanded(prev => {
      const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n
    })
  }

  const insertRef = (sheetId: string, sheetName: string, recordName: string) => {
    const ref = sheetId === currentSheetId
      ? `[${recordName}]`
      : `[${sheetName}::${recordName}]`
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

  /** Insert (axisName=recordName) as a parameter on the nearest ref before cursor.
   *  - If cursor is inside existing (...) → append ", axisName=recordName" before closing )
   *  - Else if text before cursor ends with ] → append (axisName=recordName) there
   *  - Else insert at cursor
   */
  const insertParam = useCallback((axisName: string, recordName: string) => {
    const el = textRef.current
    const param = `${axisName}=${recordName}`
    if (!el) {
      setText(prev => prev + `(${param})`)
      return
    }
    const pos = el.selectionStart
    const t = text

    // Check if cursor is inside (...) that directly follows a ]
    let parenStart = -1
    let depth = 0
    for (let i = pos - 1; i >= 0; i--) {
      if (t[i] === ')') { depth++; continue }
      if (t[i] === '(') {
        if (depth === 0) { parenStart = i; break }
        depth--
      }
    }

    if (parenStart >= 0 && parenStart > 0 && t[parenStart - 1] === ']') {
      // Inside existing parens — find closing ) at or after pos
      let closePos = pos
      while (closePos < t.length && t[closePos] !== ')') closePos++
      const inside = t.slice(parenStart + 1, closePos).trim()
      const sep = inside.length > 0 ? ', ' : ''
      const newText = t.slice(0, closePos) + sep + param + t.slice(closePos)
      setText(newText)
      const newCursor = closePos + sep.length + param.length
      setTimeout(() => { el.focus(); el.selectionStart = el.selectionEnd = newCursor }, 0)
      return
    }

    // Check if text before cursor ends with ] (possibly with whitespace)
    const before = t.slice(0, pos)
    const trimBefore = before.trimEnd()
    if (trimBefore.endsWith(']')) {
      const insertAt = trimBefore.length
      const chunk = `(${param})`
      const newText = t.slice(0, insertAt) + chunk + t.slice(insertAt)
      setText(newText)
      setTimeout(() => { el.focus(); el.selectionStart = el.selectionEnd = insertAt + chunk.length }, 0)
      return
    }

    // Fallback: insert at cursor
    const chunk = `(${param})`
    const newText = t.slice(0, pos) + chunk + t.slice(pos)
    setText(newText)
    setTimeout(() => { el.focus(); el.selectionStart = el.selectionEnd = pos + chunk.length }, 0)
  }, [text])

  const renderRecordTree = (nodes: RecordNode[], sheetId: string, sheetName: string, level: number): React.ReactNode => {
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
              onClick={() => insertRef(sheetId, sheetName, name)}
            >
              {name}
            </Typography>
          </Box>
          {hasChildren && isExp && renderRecordTree(n.children, sheetId, sheetName, level + 1)}
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
                {sExp && renderRecordTree(st.records, st.sheet.id, st.sheet.name, 0)}
                {sExp && st.axes.length > 0 && (
                  <Box sx={{ mt: 0.5 }}>
                    {st.axes.map(ax => {
                      const axKey = `axis:${st.sheet.id}:${ax.analytic.id}`
                      const axExp = expanded.has(axKey)
                      return (
                        <Box key={ax.analytic.id}>
                          <Box onClick={() => toggle(axKey)}
                            sx={{ display: 'flex', alignItems: 'center', pl: 1, py: 0.25, cursor: 'pointer', '&:hover': { bgcolor: '#f5f5f5' } }}>
                            {axExp ? <ExpandMoreOutlined sx={{ fontSize: 12, opacity: 0.5 }} /> : <ChevronRightOutlined sx={{ fontSize: 12, opacity: 0.5 }} />}
                            <Typography variant="caption" sx={{ fontSize: 10, color: '#aaa', ml: 0.5 }}>
                              {ax.analytic.name}
                            </Typography>
                          </Box>
                          {axExp && ax.records.map(n => {
                            const recName = n.data?.name || n.record.id.slice(0, 8)
                            return (
                              <Box key={n.record.id}
                                onClick={() => insertParam(ax.analytic.name, recName)}
                                sx={{ pl: 4, py: 0.25, cursor: 'pointer', fontSize: 11, color: '#555', '&:hover': { bgcolor: '#fff3e0' } }}>
                                <Typography variant="caption" sx={{ fontSize: 11 }}>
                                  {recName}
                                </Typography>
                              </Box>
                            )
                          })}
                        </Box>
                      )
                    })}
                  </Box>
                )}
              </Box>
            )
          })}
        </Box>

        {/* Center: formula editor */}
        <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
          <Typography variant="caption" sx={{ color: '#999', mb: 0.5 }}>
            Формула — клик по показателю вставляет ссылку; клик по значению оси добавляет параметр
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
            placeholder={'Пример:\nесли [Выручка] > 0 то\n  [Выручка] * 0.2\nиначе\n  0\nконец_если'}
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
