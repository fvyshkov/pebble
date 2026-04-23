import { useState, useEffect } from 'react'
import {
  TextField, Typography, Box, Switch, FormControlLabel,
  Chip, Button, IconButton, Select, MenuItem, InputLabel, FormControl,
  CircularProgress, Accordion, AccordionSummary, AccordionDetails,
} from '@mui/material'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import * as Icons from '@mui/icons-material'
import { useTranslation } from 'react-i18next'
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

export default function AnalyticSettings({ analyticId, modelId, onRefresh }: Props) {
  const { t } = useTranslation()
  const [data, setData] = useState<Analytic | null>(null)
  const [iconOpen, setIconOpen] = useState(false)
  const [sheetStatus, setSheetStatus] = useState<string>('')
  const [bulkBusy, setBulkBusy] = useState<null | 'add' | 'remove'>(null)
  const [bulkProgress, setBulkProgress] = useState<{ done: number; total: number }>({ done: 0, total: 0 })
  const [isMain, setIsMain] = useState(false)
  const { addOp, getOverrides } = usePending()

  const PERIOD_TYPES = [
    { key: 'year', label: t('analytic.year') },
    { key: 'quarter', label: t('analytic.quarter') },
    { key: 'month', label: t('analytic.month') },
  ]

  useEffect(() => {
    api.getAnalytic(analyticId).then(a => {
      const overrides = getOverrides(`analytic:${analyticId}`)
      if (overrides) {
        setData({ ...a, ...overrides, period_types: overrides.period_types ?? a.period_types })
      } else {
        setData(a)
      }
    })
    // Check if this analytic is main on any sheet.
    ;(async () => {
      try {
        const sheets = await api.listSheets(modelId)
        for (const s of sheets) {
          const m = await api.getMainAnalytic(s.id)
          if (m.analytic_id === analyticId) { setIsMain(true); return }
        }
        setIsMain(false)
      } catch { /* ignore */ }
    })()
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
    <Accordion defaultExpanded={true} disableGutters sx={{ maxWidth: 500, mb: 1, '&:before': { display: 'none' }, boxShadow: 'none', border: '1px solid #e0e0e0', borderRadius: '4px !important' }}>
      <AccordionSummary expandIcon={<ExpandMoreOutlined />}>
        <Typography variant="subtitle1" sx={{ fontWeight: 500 }}>{data.name}</Typography>
      </AccordionSummary>
      <AccordionDetails sx={{ pt: 0, px: 2, pb: 2 }}>
      <Box sx={{ display: 'flex', gap: 1, mb: 2, alignItems: 'center' }}>
        <IconButton
          onClick={() => setIconOpen(true)}
          sx={{ border: '1px solid #e0e0e0', borderRadius: 1, width: 40, height: 40 }}
        >
          {SelectedIcon ? <SelectedIcon fontSize="small" /> : <Icons.CategoryOutlined fontSize="small" />}
        </IconButton>
        <TextField
          label={t('analytic.name')} fullWidth value={data.name}
          onChange={e => {
            const name = e.target.value
            const code = transliterate(name)
            change({ name, code } as any)
          }}
        />
      </Box>

      <TextField
        label={t('analytic.code')} fullWidth value={data.code}
        onChange={e => change({ code: e.target.value } as any)}
        sx={{ mb: 2 }}
        InputProps={{ sx: { fontFamily: 'monospace' } }}
      />

      <Box sx={{ display: 'flex', gap: 3, mb: 2 }}>
        <FormControlLabel
          control={<Switch checked={isPeriods} onChange={e => change({ is_periods: e.target.checked ? 1 : 0 } as any)} />}
          label={t('analytic.periods')}
        />
        <FormControlLabel
          control={<Switch checked={isMain} onChange={async e => {
            const on = e.target.checked
            setIsMain(on)
            if (on) {
              // Set this analytic as main on all sheets it's bound to
              const sheets = await api.listSheets(modelId)
              for (const s of sheets) {
                const bindings = await api.listSheetAnalytics(s.id)
                if (bindings.some((b: any) => b.analytic_id === analyticId)) {
                  await api.setMainAnalytic(s.id, analyticId)
                }
              }
            }
            // Note: toggling off is not implemented — another analytic should be set as main instead
          }} />}
          label={t('analytic.indicators')}
        />
      </Box>

      <FormControl fullWidth sx={{ mb: 2 }}>
        <InputLabel>{t('analytic.dataType')}</InputLabel>
        <Select
          value={data.data_type || 'sum'}
          label={t('analytic.dataType')}
          onChange={e => change({ data_type: e.target.value } as any)}
        >
          <MenuItem value="sum">{t('analytic.sum')}</MenuItem>
          <MenuItem value="percent">{t('analytic.percent')}</MenuItem>
          <MenuItem value="quantity">{t('analytic.quantity')}</MenuItem>
          <MenuItem value="string">{t('analytic.string')}</MenuItem>
        </Select>
      </FormControl>

      {isPeriods && (
        <Box sx={{ ml: 1, mb: 2 }}>
          <Typography variant="body2" sx={{ mb: 1 }}>{t('analytic.periodTypes')}</Typography>
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
              label={t('analytic.start')} type="date" fullWidth
              value={data.period_start || ''}
              onChange={e => change({ period_start: e.target.value } as any)}
              InputLabelProps={{ shrink: true }}
            />
            <TextField
              label={t('analytic.end')} type="date" fullWidth
              value={data.period_end || ''}
              onChange={e => change({ period_end: e.target.value } as any)}
              InputLabelProps={{ shrink: true }}
            />
          </Box>
          <Button variant="outlined" onClick={handleGenerate} disabled={periodTypes.length === 0}>
            {t('analytic.generatePeriods')}
          </Button>
        </Box>
      )}

      {/* Add/remove from all sheets */}
      <Box sx={{ display: 'flex', gap: 1, mt: 3, mb: 1, alignItems: 'center' }}>
        <Button
          variant="outlined" size="small"
          startIcon={bulkBusy === 'add'
            ? <CircularProgress size={14} thickness={5} />
            : <Icons.PlaylistAddOutlined />}
          disabled={!!bulkBusy}
          onClick={async () => {
            setBulkBusy('add')
            try {
              const result = await api.bulkAddAnalytic(modelId, analyticId)
              const msg = result.added > 0
                ? t('analytic.addedToSheets', { count: result.added }) + (result.formulas_suggested > 0 ? t('analytic.formulasSuggested', { count: result.formulas_suggested }) : '')
                : t('analytic.alreadyInAll')
              setSheetStatus(msg)
            } catch (e: any) {
              setSheetStatus(`${t('common.error')}: ${e.message || e}`)
            } finally {
              setBulkBusy(null)
              setTimeout(() => setSheetStatus(''), 4000)
            }
          }}
          sx={{ textTransform: 'none', fontSize: 12 }}
        >
          {bulkBusy === 'add'
            ? t('analytic.adding')
            : t('analytic.addToAll')}
        </Button>
        <Button
          variant="outlined" size="small" color="warning"
          startIcon={bulkBusy === 'remove'
            ? <CircularProgress size={14} thickness={5} />
            : <Icons.PlaylistRemoveOutlined />}
          disabled={!!bulkBusy}
          onClick={async () => {
            setBulkBusy('remove')
            try {
              const result = await api.bulkRemoveAnalytic(modelId, analyticId)
              setSheetStatus(result.removed > 0 ? t('analytic.removedFromSheets', { count: result.removed }) : t('analytic.notInAny'))
            } catch (e: any) {
              setSheetStatus(`${t('common.error')}: ${e.message || e}`)
            } finally {
              setBulkBusy(null)
              setTimeout(() => setSheetStatus(''), 4000)
            }
          }}
          sx={{ textTransform: 'none', fontSize: 12 }}
        >
          {bulkBusy === 'remove'
            ? t('analytic.removing')
            : t('analytic.removeFromAll')}
        </Button>
      </Box>
      {sheetStatus && <Typography variant="caption" color="success.main">{sheetStatus}</Typography>}

      <IconPickerDialog
        open={iconOpen}
        onClose={() => setIconOpen(false)}
        onSelect={icon => change({ icon } as any)}
      />
      </AccordionDetails>
    </Accordion>
  )
}
