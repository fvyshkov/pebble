import { useState, useEffect, useMemo, useRef } from 'react'
import {
  Box, Typography, TextField, IconButton, Tooltip, Chip,
  Stack, Accordion, AccordionSummary, AccordionDetails, Button,
  ToggleButton, ToggleButtonGroup, Popover,
} from '@mui/material'
import { SimpleTreeView } from '@mui/x-tree-view/SimpleTreeView'
import { TreeItem } from '@mui/x-tree-view/TreeItem'
import AddOutlined from '@mui/icons-material/AddOutlined'
import DeleteOutlineOutlined from '@mui/icons-material/DeleteOutlineOutlined'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import DragIndicatorOutlined from '@mui/icons-material/DragIndicatorOutlined'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import FormulaEditor from './FormulaEditor'
import * as api from '../../api'
import type { Analytic, AnalyticRecord, SheetAnalytic } from '../../types'

interface Props {
  sheetId: string
  modelId: string
  indicatorId: string
  indicatorName: string
}

type Mode = 'manual' | 'formula'

interface ScopedRule {
  id?: string
  scope: Record<string, string>
  priority: number
  formula: string
  mode: Mode // UI-only; persisted as empty formula when manual
}

interface AnalyticInfo {
  id: string
  name: string
  records: AnalyticRecord[]
  byId: Record<string, AnalyticRecord>
  childrenOf: Record<string, string[]>
}

function recName(r: AnalyticRecord): string {
  try {
    const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
    return (d && (d.name || d.code)) || r.id.slice(0, 6)
  } catch {
    return r.id.slice(0, 6)
  }
}

function buildAnalyticInfo(a: Analytic, recs: AnalyticRecord[]): AnalyticInfo {
  const byId: Record<string, AnalyticRecord> = {}
  const childrenOf: Record<string, string[]> = {}
  for (const r of recs) {
    byId[r.id] = r
    const parent = r.parent_id || '__root__'
    ;(childrenOf[parent] ||= []).push(r.id)
  }
  return { id: a.id, name: a.name, records: recs, byId, childrenOf }
}

/** Tree-style record picker (Popover + SimpleTreeView). */
function RecordTreePicker({
  analytic, value, onChange,
}: {
  analytic: AnalyticInfo
  value: string
  onChange: (rid: string) => void
}) {
  const [anchor, setAnchor] = useState<HTMLElement | null>(null)
  const label = value ? recName(analytic.byId[value]) : 'любое'

  const renderNode = (id: string): React.ReactNode => {
    const rec = analytic.byId[id]
    if (!rec) return null
    const kids = analytic.childrenOf[id] || []
    return (
      <TreeItem
        key={id}
        itemId={id}
        label={recName(rec)}
        onClick={e => {
          e.stopPropagation()
          onChange(id)
          setAnchor(null)
        }}
      >
        {kids.map(renderNode)}
      </TreeItem>
    )
  }
  const roots = analytic.childrenOf['__root__'] || []

  return (
    <>
      <Button
        size="small" variant="outlined"
        onClick={e => setAnchor(e.currentTarget)}
        sx={{ textTransform: 'none', justifyContent: 'flex-start', minWidth: 140 }}
        data-testid="scope-picker-btn"
      >
        <Typography variant="caption" sx={{ fontWeight: 500, mr: 1, color: 'text.secondary' }}>
          {analytic.name}:
        </Typography>
        <Typography variant="body2" noWrap>{label}</Typography>
      </Button>
      <Popover
        open={!!anchor}
        anchorEl={anchor}
        onClose={() => setAnchor(null)}
        anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
      >
        <Box sx={{ p: 1, minWidth: 260, maxHeight: 400, overflow: 'auto' }}>
          <Button
            size="small" fullWidth
            sx={{ justifyContent: 'flex-start', textTransform: 'none', mb: 0.5 }}
            onClick={() => { onChange(''); setAnchor(null) }}
          >
            <em>любое</em>
          </Button>
          <SimpleTreeView
            defaultExpandedItems={roots}
            sx={{ '& .MuiTreeItem-content': { py: 0.25 } }}
          >
            {roots.map(renderNode)}
          </SimpleTreeView>
        </Box>
      </Popover>
    </>
  )
}

export default function IndicatorFormulasPanel({
  sheetId, modelId, indicatorId, indicatorName,
}: Props) {
  const [leafFormula, setLeafFormula] = useState('')
  const [leafMode, setLeafMode] = useState<Mode>('manual')
  const [consolFormula, setConsolFormula] = useState('')
  const [consolMode, setConsolMode] = useState<Mode>('manual')
  const [scoped, setScoped] = useState<ScopedRule[]>([])
  const [loading, setLoading] = useState(false)
  const [saving, setSaving] = useState(false)
  const [dirty, setDirty] = useState(false)
  const [analytics, setAnalytics] = useState<AnalyticInfo[]>([])
  const [mainAid, setMainAid] = useState<string | null>(null)
  const [expanded, setExpanded] = useState<Record<string, boolean>>({ consol: true, leaf: true })
  // Popup editor slot (which formula is being edited in FormulaEditor).
  const [editorSlot, setEditorSlot] = useState<
    | { kind: 'consol' }
    | { kind: 'leaf' }
    | { kind: 'scoped'; idx: number }
    | null
  >(null)
  const editorFormula = editorSlot
    ? (editorSlot.kind === 'consol'
        ? consolFormula
        : editorSlot.kind === 'leaf'
          ? leafFormula
          : scoped[editorSlot.idx]?.formula || '')
    : ''

  const dragIdxRef = useRef<number | null>(null)

  useEffect(() => {
    if (!sheetId || !indicatorId) return
    setLoading(true)
    setDirty(false)
    ;(async () => {
      try {
        const [rules, main, bindings] = await Promise.all([
          api.getIndicatorRules(sheetId, indicatorId),
          api.getMainAnalytic(sheetId),
          api.listSheetAnalytics(sheetId),
        ])
        setConsolFormula(rules.consolidation || '')
        setConsolMode(rules.consolidation ? 'formula' : 'manual')
        setLeafFormula(rules.leaf || '')
        setLeafMode(rules.leaf ? 'formula' : 'manual')
        // Top = highest priority
        const s = (rules.scoped || [])
          .slice()
          .sort((a, b) => b.priority - a.priority)
          .map(r => ({ ...r, mode: (r.formula ? 'formula' : 'manual') as Mode }))
        setScoped(s)
        setMainAid(main.analytic_id)
        const nonMain = bindings.filter((b: SheetAnalytic) => b.analytic_id !== main.analytic_id)
        const infos: AnalyticInfo[] = []
        for (const b of nonMain) {
          const [a, recs] = await Promise.all([
            api.getAnalytic(b.analytic_id),
            api.listRecords(b.analytic_id),
          ])
          infos.push(buildAnalyticInfo(a as Analytic, recs))
        }
        setAnalytics(infos)
      } finally {
        setLoading(false)
      }
    })()
  }, [sheetId, indicatorId])

  const anameById = useMemo(() => {
    const m: Record<string, string> = {}
    for (const a of analytics) m[a.id] = a.name
    return m
  }, [analytics])

  const recLabel = (aid: string, rid: string): string => {
    const a = analytics.find(x => x.id === aid)
    if (!a) return rid.slice(0, 4)
    const r = a.byId[rid]
    return r ? recName(r) : rid.slice(0, 4)
  }

  const markDirty = () => setDirty(true)
  const toggleExpanded = (key: string) => {
    setExpanded(prev => ({ ...prev, [key]: !prev[key] }))
  }

  const handleAddScoped = () => {
    // New rules are formula-mode by default (a rule exists because you want a formula).
    const maxPrio = scoped.reduce((m, r) => Math.max(m, r.priority), 100)
    const newRule: ScopedRule = { scope: {}, priority: maxPrio + 10, formula: '', mode: 'formula' }
    setScoped(prev => [newRule, ...prev])
    setExpanded(prev => ({ ...prev, 'scoped-0': true }))
    markDirty()
  }
  const handleDeleteScoped = (idx: number) => {
    setScoped(prev => prev.filter((_, i) => i !== idx))
    markDirty()
  }
  const handleScopeChange = (idx: number, aid: string, rid: string) => {
    setScoped(prev => prev.map((r, i) => {
      if (i !== idx) return r
      const next = { ...r.scope }
      if (!rid) delete next[aid]
      else next[aid] = rid
      return { ...r, scope: next }
    }))
    markDirty()
  }
  const patchScoped = (idx: number, patch: Partial<ScopedRule>) => {
    setScoped(prev => prev.map((r, i) => i === idx ? { ...r, ...patch } : r))
    markDirty()
  }

  const handleDragStart = (idx: number) => (e: React.DragEvent) => {
    dragIdxRef.current = idx
    e.dataTransfer.effectAllowed = 'move'
  }
  const handleDragOver = (e: React.DragEvent) => { e.preventDefault() }
  const handleDrop = (idx: number) => (e: React.DragEvent) => {
    e.preventDefault()
    const from = dragIdxRef.current
    dragIdxRef.current = null
    if (from == null || from === idx) return
    setScoped(prev => {
      const next = prev.slice()
      const [moved] = next.splice(from, 1)
      next.splice(idx, 0, moved)
      const N = next.length
      return next.map((r, i) => ({ ...r, priority: 100 + (N - i) * 10 }))
    })
    markDirty()
  }

  const handleSave = async () => {
    setSaving(true)
    try {
      await api.putIndicatorRules(sheetId, indicatorId, {
        // Manual mode = persist empty formula.
        consolidation: consolMode === 'formula' ? consolFormula : '',
        leaf: leafMode === 'formula' ? leafFormula : '',
        scoped: scoped.map(r => ({
          id: r.id,
          scope: r.scope,
          priority: r.priority,
          formula: r.mode === 'formula' ? r.formula : '',
        })),
      })
      setDirty(false)
    } finally {
      setSaving(false)
    }
  }

  if (loading) {
    return (
      <Box sx={{ p: 2 }}>
        <Typography variant="body2" color="text.secondary">Загрузка…</Typography>
      </Box>
    )
  }

  // Sub-component: mode toggle + conditional formula field.
  const SlotBody = ({
    mode, onModeChange, formula, onFormulaChange, onEdit, placeholder, hint, manualHint,
  }: {
    mode: Mode
    onModeChange: (m: Mode) => void
    formula: string
    onFormulaChange: (v: string) => void
    onEdit?: () => void
    placeholder?: string
    hint?: string
    manualHint?: string
  }) => (
    <Box>
      <ToggleButtonGroup
        exclusive size="small" value={mode}
        onChange={(_, v) => { if (v) onModeChange(v) }}
        sx={{ mb: 1 }}
      >
        <ToggleButton value="manual" sx={{ textTransform: 'none', py: 0.25, px: 1 }}>
          Ручной ввод
        </ToggleButton>
        <ToggleButton value="formula" sx={{ textTransform: 'none', py: 0.25, px: 1 }}>
          Формула
        </ToggleButton>
      </ToggleButtonGroup>
      {mode === 'formula' ? (
        <>
          <Stack direction="row" spacing={1} alignItems="flex-start">
            <TextField
              multiline minRows={1} maxRows={4} fullWidth size="small"
              value={formula}
              onChange={e => onFormulaChange(e.target.value)}
              onClick={onEdit}
              placeholder={placeholder}
              InputProps={{ sx: { fontFamily: 'monospace', fontSize: 13 } }}
            />
            {onEdit && (
              <Tooltip title="Открыть редактор формул">
                <IconButton size="small" onClick={onEdit}>
                  <EditOutlined fontSize="small" />
                </IconButton>
              </Tooltip>
            )}
          </Stack>
          {hint && (
            <Typography variant="caption" color="text.secondary">{hint}</Typography>
          )}
        </>
      ) : (
        <Typography variant="caption" color="text.secondary">
          {manualHint || 'Значение вводится пользователем в клетке.'}
        </Typography>
      )}
    </Box>
  )

  return (
    <Box sx={{ p: 1 }} data-testid="indicator-formulas-panel">
      <Stack direction="row" alignItems="center" spacing={1} sx={{ mb: 1 }}>
        <Typography variant="subtitle2" noWrap sx={{ flex: 1, minWidth: 0 }}>
          Формулы: {indicatorName}
        </Typography>
        <Tooltip title="Добавить правило">
          <IconButton size="small" onClick={handleAddScoped}>
            <AddOutlined fontSize="small" />
          </IconButton>
        </Tooltip>
        <Button
          size="small" variant="contained" startIcon={<SaveOutlined />}
          disabled={!dirty || saving}
          onClick={handleSave}
        >Сохранить</Button>
      </Stack>

      {mainAid == null && (
        <Typography variant="caption" color="warning.main" sx={{ display: 'block', mb: 1 }}>
          Главная аналитика листа не задана — правила не применятся.
        </Typography>
      )}

      {/* Слот 1: Консолидация (top) */}
      <Accordion
        expanded={!!expanded['consol']}
        onChange={() => toggleExpanded('consol')}
        disableGutters
        data-testid="formula-slot-consol"
      >
        <AccordionSummary expandIcon={<ExpandMoreOutlined fontSize="small" />} sx={{ minHeight: 36, '& .MuiAccordionSummary-content': { my: 0.5 } }}>
          <Stack direction="row" spacing={1} alignItems="center" sx={{ flex: 1, minWidth: 0 }}>
            <Chip size="small" color="primary" variant="outlined" label="консолидация" />
            <Chip size="small" label={consolMode === 'formula' ? 'формула' : 'ручной ввод'} />
            <Typography variant="caption" color="text.secondary" noWrap sx={{ flex: 1, minWidth: 0, fontFamily: 'monospace' }}>
              {consolMode === 'formula' ? (consolFormula || '—') : 'сумма нижестоящих (по умолчанию)'}
            </Typography>
          </Stack>
        </AccordionSummary>
        <AccordionDetails sx={{ pt: 0 }}>
          <SlotBody
            mode={consolMode}
            onModeChange={m => { setConsolMode(m); markDirty() }}
            formula={consolFormula}
            onFormulaChange={v => { setConsolFormula(v); markDirty() }}
            onEdit={() => setEditorSlot({ kind: 'consol' })}
            placeholder="например: [выдачи] / [партнёры]"
            hint="Применяется на HEAD-клетках (когда хотя бы одна не-главная ось — не лист)."
            manualHint="По умолчанию — сумма нижестоящих. Переключите в «Формула» для агрегата через выражение."
          />
        </AccordionDetails>
      </Accordion>

      {/* Слоты 2..N: scoped (draggable) */}
      {scoped.map((r, idx) => {
        const key = `scoped-${idx}`
        return (
          <Accordion
            key={r.id || `new-${idx}`}
            expanded={!!expanded[key]}
            onChange={() => toggleExpanded(key)}
            disableGutters
            data-testid="formula-slot-scoped"
            draggable
            onDragStart={handleDragStart(idx)}
            onDragOver={handleDragOver}
            onDrop={handleDrop(idx)}
          >
            <AccordionSummary
              expandIcon={<ExpandMoreOutlined fontSize="small" />}
              sx={{ minHeight: 36, '& .MuiAccordionSummary-content': { my: 0.5 } }}
            >
              <Stack direction="row" spacing={1} alignItems="center" sx={{ flex: 1, minWidth: 0 }}>
                <DragIndicatorOutlined fontSize="small" sx={{ color: 'text.disabled', cursor: 'grab' }} />
                <Chip size="small" label={`prio ${r.priority}`} />
                <Chip size="small" color={r.mode === 'formula' ? 'primary' : 'default'}
                  label={r.mode === 'formula' ? 'формула' : 'ручной ввод'} />
                {Object.entries(r.scope).slice(0, 2).map(([aid, rid]) => (
                  <Chip key={aid} size="small" variant="outlined"
                    label={`${anameById[aid] || aid.slice(0, 4)}: ${recLabel(aid, rid)}`} />
                ))}
                {Object.keys(r.scope).length > 2 && (
                  <Chip size="small" variant="outlined" label={`+${Object.keys(r.scope).length - 2}`} />
                )}
                <Typography variant="caption" color="text.secondary" noWrap sx={{ flex: 1, minWidth: 0, fontFamily: 'monospace' }}>
                  {r.mode === 'formula' ? (r.formula || '—') : '(ручной ввод)'}
                </Typography>
              </Stack>
            </AccordionSummary>
            <AccordionDetails sx={{ pt: 0 }}>
              <Stack direction="row" spacing={1} alignItems="center" sx={{ mb: 1, flexWrap: 'wrap' }}>
                <Typography variant="caption" sx={{ color: 'text.secondary' }}>Область:</Typography>
                {analytics.map(a => (
                  <RecordTreePicker
                    key={a.id}
                    analytic={a}
                    value={r.scope[a.id] || ''}
                    onChange={rid => handleScopeChange(idx, a.id, rid)}
                  />
                ))}
                <TextField
                  size="small" type="number" label="Приоритет"
                  value={r.priority}
                  onChange={e => patchScoped(idx, { priority: Number(e.target.value || 0) })}
                  sx={{ width: 100 }}
                />
                <Box sx={{ flex: 1 }} />
                <Tooltip title="Удалить правило">
                  <IconButton size="small" onClick={() => handleDeleteScoped(idx)}>
                    <DeleteOutlineOutlined fontSize="small" />
                  </IconButton>
                </Tooltip>
              </Stack>
              <SlotBody
                mode={r.mode}
                onModeChange={m => patchScoped(idx, { mode: m })}
                formula={r.formula}
                onFormulaChange={v => patchScoped(idx, { formula: v })}
                onEdit={() => setEditorSlot({ kind: 'scoped', idx })}
                placeholder="формула"
              />
            </AccordionDetails>
          </Accordion>
        )
      })}

      {/* Последний слот: Обычная клетка (bottom) */}
      <Accordion
        expanded={!!expanded['leaf']}
        onChange={() => toggleExpanded('leaf')}
        disableGutters
        data-testid="formula-slot-leaf"
      >
        <AccordionSummary expandIcon={<ExpandMoreOutlined fontSize="small" />} sx={{ minHeight: 36, '& .MuiAccordionSummary-content': { my: 0.5 } }}>
          <Stack direction="row" spacing={1} alignItems="center" sx={{ flex: 1, minWidth: 0 }}>
            <Chip size="small" variant="outlined" label="обычная клетка" />
            <Chip size="small" label={leafMode === 'formula' ? 'формула' : 'ручной ввод'} />
            <Typography variant="caption" color="text.secondary" noWrap sx={{ flex: 1, minWidth: 0, fontFamily: 'monospace' }}>
              {leafMode === 'formula' ? (leafFormula || '—') : 'значение вводит пользователь'}
            </Typography>
          </Stack>
        </AccordionSummary>
        <AccordionDetails sx={{ pt: 0 }}>
          <SlotBody
            mode={leafMode}
            onModeChange={m => { setLeafMode(m); markDirty() }}
            formula={leafFormula}
            onFormulaChange={v => { setLeafFormula(v); markDirty() }}
            onEdit={() => setEditorSlot({ kind: 'leaf' })}
            placeholder="например: [выдачи] * 0.1"
            hint="База для листовой клетки (все не-главные оси — листья)."
            manualHint="По умолчанию — ручной ввод значения. Переключите в «Формула» для вычисляемого показателя."
          />
        </AccordionDetails>
      </Accordion>

      {/* Popup editor — reuses the same tree-of-analytics picker. */}
      <FormulaEditor
        open={!!editorSlot}
        formula={editorFormula}
        modelId={modelId}
        onClose={() => setEditorSlot(null)}
        onSave={text => {
          if (!editorSlot) return
          if (editorSlot.kind === 'consol') { setConsolFormula(text); setConsolMode('formula') }
          else if (editorSlot.kind === 'leaf') { setLeafFormula(text); setLeafMode('formula') }
          else patchScoped(editorSlot.idx, { formula: text, mode: 'formula' })
          markDirty()
        }}
      />
    </Box>
  )
}
