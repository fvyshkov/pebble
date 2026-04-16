import { useState, useEffect } from 'react'
import {
  TextField, Typography, Box, Switch, FormControlLabel,
  Chip, Button, IconButton, Select, MenuItem, InputLabel, FormControl,
} from '@mui/material'
import * as Icons from '@mui/icons-material'
import * as api from '../../api'
import type { Analytic } from '../../types'
import { transliterate } from '../../utils/transliterate'
import IconPickerDialog from '../../components/IconPickerDialog'
import { usePending } from '../../store/PendingContext'

interface Props {
  analyticId: string
  modelId: string
  onRefresh: () => void
}

const PERIOD_TYPES = [
  { key: 'year', label: 'Год' },
  { key: 'quarter', label: 'Квартал' },
  { key: 'month', label: 'Месяц' },
]

export default function AnalyticSettings({ analyticId, modelId, onRefresh }: Props) {
  const [data, setData] = useState<Analytic | null>(null)
  const [iconOpen, setIconOpen] = useState(false)
  const [sheetStatus, setSheetStatus] = useState<string>('')
  const { addOp, getOverrides } = usePending()

  useEffect(() => {
    api.getAnalytic(analyticId).then(a => {
      const overrides = getOverrides(`analytic:${analyticId}`)
      if (overrides) {
        setData({ ...a, ...overrides, period_types: overrides.period_types ?? a.period_types })
      } else {
        setData(a)
      }
    })
  }, [analyticId])

  if (!data) return null

  const isPeriods = !!data.is_periods
  const periodTypes: string[] = (() => {
    try {
      const pt = data.period_types
      return typeof pt === 'string' ? JSON.parse(pt) : pt
    } catch { return [] }
  })()

  const change = (updates: Partial<Analytic>) => {
    const updated = { ...data, ...updates } as Analytic
    setData(updated)
    const ptParsed = (() => {
      try {
        const pt = updated.period_types
        return typeof pt === 'string' ? JSON.parse(pt) : pt
      } catch { return [] }
    })()
    addOp({
      key: `analytic:${analyticId}`,
      type: 'updateAnalytic',
      id: analyticId,
      data: {
        name: updated.name,
        code: updated.code,
        icon: updated.icon,
        is_periods: !!updated.is_periods,
        data_type: updated.data_type || 'sum',
        period_types: ptParsed,
        period_start: updated.period_start,
        period_end: updated.period_end,
        sort_order: updated.sort_order,
      },
    })
  }

  const togglePeriodType = (type: string) => {
    const next = periodTypes.includes(type) ? periodTypes.filter(t => t !== type) : [...periodTypes, type]
    change({ period_types: JSON.stringify(next) } as any)
  }

  const handleGenerate = async () => {
    // Flush this analytic's settings to backend first
    const ptParsed = (() => {
      try {
        const pt = data.period_types
        return typeof pt === 'string' ? JSON.parse(pt) : pt
      } catch { return [] }
    })()
    await api.updateAnalytic(analyticId, {
      name: data.name,
      code: data.code,
      icon: data.icon,
      is_periods: !!data.is_periods,
      data_type: data.data_type || 'sum',
      period_types: ptParsed,
      period_start: data.period_start,
      period_end: data.period_end,
      sort_order: data.sort_order,
    })
    await api.generatePeriods(analyticId)
    onRefresh()
  }

  const SelectedIcon = data.icon ? (Icons as any)[data.icon] : null

  return (
    <Box sx={{ maxWidth: 500, mb: 3 }}>
      <Typography variant="h6" sx={{ mb: 2 }}>Аналитика</Typography>

      <Box sx={{ display: 'flex', gap: 1, mb: 2, alignItems: 'center' }}>
        <IconButton
          onClick={() => setIconOpen(true)}
          sx={{ border: '1px solid #e0e0e0', borderRadius: 1, width: 40, height: 40 }}
        >
          {SelectedIcon ? <SelectedIcon fontSize="small" /> : <Icons.CategoryOutlined fontSize="small" />}
        </IconButton>
        <TextField
          label="Название" fullWidth value={data.name}
          onChange={e => {
            const name = e.target.value
            const code = transliterate(name)
            change({ name, code } as any)
          }}
        />
      </Box>

      <TextField
        label="Код" fullWidth value={data.code}
        onChange={e => change({ code: e.target.value } as any)}
        sx={{ mb: 2 }}
        InputProps={{ sx: { fontFamily: 'monospace' } }}
      />

      <FormControlLabel
        control={<Switch checked={isPeriods} onChange={e => change({ is_periods: e.target.checked ? 1 : 0 } as any)} />}
        label="Периоды"
        sx={{ mb: 2 }}
      />

      <FormControl fullWidth sx={{ mb: 2 }}>
        <InputLabel>Тип данных</InputLabel>
        <Select
          value={data.data_type || 'sum'}
          label="Тип данных"
          onChange={e => change({ data_type: e.target.value } as any)}
        >
          <MenuItem value="sum">Сумма (0.00)</MenuItem>
          <MenuItem value="percent">% (процент)</MenuItem>
          <MenuItem value="quantity">Количество</MenuItem>
          <MenuItem value="string">Строка</MenuItem>
        </Select>
      </FormControl>

      {isPeriods && (
        <Box sx={{ ml: 1, mb: 2 }}>
          <Typography variant="body2" sx={{ mb: 1 }}>Типы периодов:</Typography>
          <Box sx={{ display: 'flex', gap: 1, mb: 2 }}>
            {PERIOD_TYPES.map(pt => (
              <Chip
                key={pt.key}
                label={pt.label}
                color={periodTypes.includes(pt.key) ? 'primary' : 'default'}
                variant={periodTypes.includes(pt.key) ? 'filled' : 'outlined'}
                onClick={() => togglePeriodType(pt.key)}
              />
            ))}
          </Box>
          <Box sx={{ display: 'flex', gap: 2, mb: 2 }}>
            <TextField
              label="Начало" type="date" fullWidth
              value={data.period_start || ''}
              onChange={e => change({ period_start: e.target.value } as any)}
              InputLabelProps={{ shrink: true }}
            />
            <TextField
              label="Окончание" type="date" fullWidth
              value={data.period_end || ''}
              onChange={e => change({ period_end: e.target.value } as any)}
              InputLabelProps={{ shrink: true }}
            />
          </Box>
          <Button variant="outlined" onClick={handleGenerate} disabled={periodTypes.length === 0}>
            Сгенерировать периоды
          </Button>
        </Box>
      )}

      {/* Add/remove from all sheets */}
      <Box sx={{ display: 'flex', gap: 1, mt: 3, mb: 1 }}>
        <Button
          variant="outlined" size="small"
          startIcon={<Icons.PlaylistAddOutlined />}
          onClick={async () => {
            const sheets = await api.listSheets(modelId)
            let added = 0
            for (const sheet of sheets) {
              const existing = await api.listSheetAnalytics(sheet.id)
              if (!existing.some((sa: any) => sa.analytic_id === analyticId)) {
                await api.addSheetAnalytic(sheet.id, { analytic_id: analyticId, sort_order: existing.length })
                added++
              }
            }
            setSheetStatus(added > 0 ? `Добавлено в ${added} лист(ов)` : 'Уже во всех листах')
            setTimeout(() => setSheetStatus(''), 3000)
          }}
          sx={{ textTransform: 'none', fontSize: 12 }}
        >
          Добавить во все листы
        </Button>
        <Button
          variant="outlined" size="small" color="warning"
          startIcon={<Icons.PlaylistRemoveOutlined />}
          onClick={async () => {
            const sheets = await api.listSheets(modelId)
            let removed = 0
            for (const sheet of sheets) {
              const existing = await api.listSheetAnalytics(sheet.id)
              const sa = existing.find((s: any) => s.analytic_id === analyticId)
              if (sa) {
                await api.removeSheetAnalytic(sheet.id, sa.id)
                removed++
              }
            }
            setSheetStatus(removed > 0 ? `Удалено из ${removed} лист(ов)` : 'Нет ни в одном листе')
            setTimeout(() => setSheetStatus(''), 3000)
          }}
          sx={{ textTransform: 'none', fontSize: 12 }}
        >
          Удалить со всех листов
        </Button>
      </Box>
      {sheetStatus && <Typography variant="caption" color="success.main">{sheetStatus}</Typography>}

      <IconPickerDialog
        open={iconOpen}
        onClose={() => setIconOpen(false)}
        onSelect={icon => change({ icon } as any)}
      />
    </Box>
  )
}
