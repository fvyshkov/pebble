import { useState, useEffect, useMemo } from 'react'
import {
  Dialog, DialogTitle, DialogContent, DialogActions, Button,
  Box, Typography, TextField, IconButton, Tooltip, Chip, Divider,
  MenuItem, Select, FormControl, InputLabel, Stack,
} from '@mui/material'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import * as api from '../../api'
import type { Analytic, AnalyticRecord, SheetAnalytic } from '../../types'

interface Props {
  open: boolean
  onClose: () => void
  sheetId: string
  indicatorId: string
  indicatorName: string
}

interface ScopedRule {
  id?: string
  scope: Record<string, string>
  priority: number
  formula: string
}

interface AnalyticInfo {
  id: string
  name: string
  records: { id: string; name: string }[]
}

function recName(r: AnalyticRecord): string {
  try {
    const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
    return (d && d.name) || r.id.slice(0, 6)
  } catch {
    return r.id.slice(0, 6)
  }
}

export default function IndicatorFormulasPanel({
  open, onClose, sheetId, indicatorId, indicatorName,
}: Props) {
  const [leaf, setLeaf] = useState('')
  const [consolidation, setConsolidation] = useState('')
  const [scoped, setScoped] = useState<ScopedRule[]>([])
  const [loading, setLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  // Non-main analytics bound to the sheet (for scope editor)
  const [analytics, setAnalytics] = useState<AnalyticInfo[]>([])
  const [mainAid, setMainAid] = useState<string | null>(null)

  useEffect(() => {
    if (!open || !sheetId || !indicatorId) return
    setLoading(true)
    ;(async () => {
      try {
        const [rules, main, bindings] = await Promise.all([
          api.getIndicatorRules(sheetId, indicatorId),
          api.getMainAnalytic(sheetId),
          api.listSheetAnalytics(sheetId),
        ])
        setLeaf(rules.leaf || '')
        setConsolidation(rules.consolidation || '')
        setScoped(rules.scoped || [])
        setMainAid(main.analytic_id)
        // Load non-main analytics' records for scope chips
        const nonMain = bindings.filter((b: SheetAnalytic) => b.analytic_id !== main.analytic_id)
        const infos: AnalyticInfo[] = []
        for (const b of nonMain) {
          const [a, recs] = await Promise.all([
            api.getAnalytic(b.analytic_id),
            api.listRecords(b.analytic_id),
          ])
          infos.push({
            id: a.id,
            name: (a as Analytic).name,
            records: recs.map(r => ({ id: r.id, name: recName(r) })),
          })
        }
        setAnalytics(infos)
      } finally {
        setLoading(false)
      }
    })()
  }, [open, sheetId, indicatorId])

  const anameById = useMemo(() => {
    const m: Record<string, string> = {}
    for (const a of analytics) m[a.id] = a.name
    return m
  }, [analytics])
  const rnameByAidRid = useMemo(() => {
    const m: Record<string, Record<string, string>> = {}
    for (const a of analytics) {
      m[a.id] = {}
      for (const r of a.records) m[a.id][r.id] = r.name
    }
    return m
  }, [analytics])

  const handleAddScoped = () => {
    setScoped(prev => [...prev, { scope: {}, priority: 100, formula: '' }])
  }

  const handleDeleteScoped = (idx: number) => {
    setScoped(prev => prev.filter((_, i) => i !== idx))
  }

  const handleScopeChange = (idx: number, aid: string, rid: string) => {
    setScoped(prev => prev.map((r, i) => {
      if (i !== idx) return r
      const next = { ...r.scope }
      if (!rid) delete next[aid]
      else next[aid] = rid
      return { ...r, scope: next }
    }))
  }

  const handleSave = async () => {
    setSaving(true)
    try {
      await api.putIndicatorRules(sheetId, indicatorId, {
        leaf, consolidation, scoped,
      })
      onClose()
    } finally {
      setSaving(false)
    }
  }

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle>Формулы показателя: {indicatorName}</DialogTitle>
      <DialogContent dividers>
        {loading ? (
          <Typography variant="body2" color="text.secondary">Загрузка…</Typography>
        ) : (
          <Stack spacing={2}>
            {mainAid == null && (
              <Typography variant="body2" color="warning.main">
                Главная аналитика листа не задана — правила не применятся.
              </Typography>
            )}

            <Box>
              <Typography variant="subtitle2" sx={{ mb: 0.5 }}>Обычная клетка (leaf)</Typography>
              <TextField
                multiline minRows={1} maxRows={4} fullWidth size="small"
                value={leaf} onChange={e => setLeaf(e.target.value)}
                placeholder="например:  [выдачи] * 0.1"
                InputProps={{ sx: { fontFamily: 'monospace', fontSize: 13 } }}
              />
            </Box>

            <Box>
              <Typography variant="subtitle2" sx={{ mb: 0.5 }}>Консолидация</Typography>
              <TextField
                multiline minRows={1} maxRows={4} fullWidth size="small"
                value={consolidation} onChange={e => setConsolidation(e.target.value)}
                placeholder="например:  [выдачи] / [партнёры]"
                InputProps={{ sx: { fontFamily: 'monospace', fontSize: 13 } }}
              />
              <Typography variant="caption" color="text.secondary">
                Применяется на HEAD-клетках (когда хотя бы одна не-главная ось — не лист).
                Ссылка <code>[Y]</code> резолвится в координате текущей клетки.
              </Typography>
            </Box>

            <Divider />

            <Box>
              <Stack direction="row" alignItems="center" spacing={1} sx={{ mb: 1 }}>
                <Typography variant="subtitle2">Специальные правила (scoped)</Typography>
                <Box sx={{ flex: 1 }} />
                <Button size="small" startIcon={<AddOutlined />} onClick={handleAddScoped}>
                  Добавить правило
                </Button>
              </Stack>

              {scoped.length === 0 && (
                <Typography variant="caption" color="text.secondary">
                  Правил нет.
                </Typography>
              )}

              <Stack spacing={1.5}>
                {scoped.map((r, idx) => (
                  <Box key={r.id || idx} sx={{ border: 1, borderColor: 'divider', borderRadius: 1, p: 1 }}>
                    <Stack direction="row" alignItems="flex-start" spacing={1}>
                      <Box sx={{ flex: 1 }}>
                        <Stack direction="row" spacing={1} sx={{ mb: 1, flexWrap: 'wrap' }}>
                          {analytics.map(a => (
                            <FormControl size="small" key={a.id} sx={{ minWidth: 140 }}>
                              <InputLabel>{a.name}</InputLabel>
                              <Select
                                label={a.name}
                                value={r.scope[a.id] || ''}
                                onChange={e => handleScopeChange(idx, a.id, String(e.target.value || ''))}
                              >
                                <MenuItem value=""><em>любое</em></MenuItem>
                                {a.records.map(rec => (
                                  <MenuItem key={rec.id} value={rec.id}>{rec.name}</MenuItem>
                                ))}
                              </Select>
                            </FormControl>
                          ))}
                          <TextField
                            size="small" type="number" label="Приоритет"
                            value={r.priority}
                            onChange={e => setScoped(prev => prev.map((x, i) => i === idx ? { ...x, priority: Number(e.target.value || 0) } : x))}
                            sx={{ width: 110 }}
                          />
                        </Stack>

                        <Stack direction="row" spacing={0.5} sx={{ mb: 0.5, flexWrap: 'wrap' }}>
                          {Object.entries(r.scope).map(([aid, rid]) => (
                            <Chip key={aid} size="small"
                              label={`${anameById[aid] || aid.slice(0, 4)}: ${rnameByAidRid[aid]?.[rid] || rid.slice(0, 4)}`} />
                          ))}
                          {Object.keys(r.scope).length === 0 && (
                            <Typography variant="caption" color="text.secondary">(без ограничений)</Typography>
                          )}
                        </Stack>

                        <TextField
                          multiline minRows={1} maxRows={4} fullWidth size="small"
                          value={r.formula}
                          onChange={e => setScoped(prev => prev.map((x, i) => i === idx ? { ...x, formula: e.target.value } : x))}
                          placeholder="формула"
                          InputProps={{ sx: { fontFamily: 'monospace', fontSize: 13 } }}
                        />
                      </Box>
                      <Tooltip title="Удалить правило">
                        <IconButton size="small" onClick={() => handleDeleteScoped(idx)}>
                          <DeleteOutlineOutlined fontSize="small" />
                        </IconButton>
                      </Tooltip>
                    </Stack>
                  </Box>
                ))}
              </Stack>
            </Box>
          </Stack>
        )}
      </DialogContent>
      <DialogActions>
        <Button onClick={onClose}>Отмена</Button>
        <Button variant="contained" onClick={handleSave} disabled={saving || loading}>
          Сохранить
        </Button>
      </DialogActions>
    </Dialog>
  )
}
