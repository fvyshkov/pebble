import { useState, useEffect, useMemo, useCallback } from 'react'
import {
  Box, Typography, Checkbox, IconButton, Collapse, Chip,
} from '@mui/material'
import ExpandMoreOutlined from '@mui/icons-material/ExpandMoreOutlined'
import ChevronRightOutlined from '@mui/icons-material/ChevronRightOutlined'
import { useTranslation } from 'react-i18next'
import * as api from '../../api'
import type { AnalyticRecord } from '../../types'

interface Props {
  sheetId: string
  saId: string
  analyticId: string
  isPeriods: boolean
  initialVisible: string | null  // JSON array or null (= all)
  onSaved: () => void
}

interface RecNode {
  id: string
  name: string
  periodKey: string  // empty for non-period analytics
  level: string
  children: RecNode[]
}

function periodLevel(pk: string): string {
  if (/^\d{4}-Y$/.test(pk)) return 'Y'
  if (/^\d{4}-H\d$/.test(pk)) return 'H'
  if (/^\d{4}-Q\d$/.test(pk)) return 'Q'
  if (/^\d{4}-\d{2}$/.test(pk) || /^\d{4}-M\d{2}$/.test(pk)) return 'M'
  return ''
}

function buildTree(records: AnalyticRecord[]): RecNode[] {
  const nodes = new Map<string, RecNode>()
  for (const r of records) {
    const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
    const pk = d?.period_key || ''
    nodes.set(r.id, {
      id: r.id,
      name: d?.name || pk || r.id.slice(0, 8),
      periodKey: pk,
      level: periodLevel(pk),
      children: [],
    })
  }
  const roots: RecNode[] = []
  for (const r of records) {
    const node = nodes.get(r.id)!
    if (r.parent_id && nodes.has(r.parent_id)) {
      nodes.get(r.parent_id)!.children.push(node)
    } else {
      roots.push(node)
    }
  }
  return roots
}

function getAllIds(node: RecNode): string[] {
  const ids = [node.id]
  for (const c of node.children) ids.push(...getAllIds(c))
  return ids
}

function getAllTreeIds(nodes: RecNode[]): string[] {
  return nodes.flatMap(n => getAllIds(n))
}

export default function RecordOrganizer({ sheetId, saId, analyticId, isPeriods, initialVisible, onSaved }: Props) {
  const { t } = useTranslation()
  const [records, setRecords] = useState<AnalyticRecord[]>([])
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [expanded, setExpanded] = useState<Set<string>>(new Set())
  const [loaded, setLoaded] = useState(false)

  useEffect(() => {
    api.listRecords(analyticId).then(recs => {
      setRecords(recs)
      if (initialVisible) {
        try {
          setSelected(new Set(JSON.parse(initialVisible)))
        } catch {
          setSelected(new Set(recs.map(r => r.id)))
        }
      } else {
        setSelected(new Set(recs.map(r => r.id)))
      }
      const tree = buildTree(recs)
      setExpanded(new Set(tree.map(n => n.id)))
      setLoaded(true)
    })
  }, [analyticId, initialVisible])

  const tree = useMemo(() => buildTree(records), [records])
  const allIds = useMemo(() => getAllTreeIds(tree), [tree])

  const save = useCallback(async (newSelected: Set<string>) => {
    const visibleIds = newSelected.size === allIds.length ? null : Array.from(newSelected)
    await api.setPeriodLevel(sheetId, saId, null, visibleIds)
    onSaved()
  }, [sheetId, saId, allIds.length, onSaved])

  const toggleNode = useCallback((node: RecNode) => {
    setSelected(prev => {
      const next = new Set(prev)
      const ids = getAllIds(node)
      const allChecked = ids.every(id => prev.has(id))
      for (const id of ids) allChecked ? next.delete(id) : next.add(id)
      save(next)
      return next
    })
  }, [save])

  const toggleExpand = useCallback((id: string) => {
    setExpanded(prev => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }, [])

  const selectAll = useCallback(() => {
    const all = new Set(allIds)
    setSelected(all)
    save(all)
  }, [allIds, save])

  const selectNone = useCallback(() => {
    const s = new Set<string>()
    setSelected(s)
    save(s)
  }, [save])

  const selectByLevel = useCallback((level: string) => {
    const ranks: Record<string, number> = { M: 0, Q: 1, H: 2, Y: 3 }
    const minRank = ranks[level] ?? 0
    const newSel = new Set<string>()
    for (const r of records) {
      const d = typeof r.data_json === 'string' ? JSON.parse(r.data_json) : r.data_json
      const pk = d?.period_key || ''
      const lvl = periodLevel(pk)
      const rank = ranks[lvl] ?? -1
      if (rank >= minRank) newSel.add(r.id)
    }
    setSelected(newSel)
    save(newSel)
  }, [records, save])

  if (!loaded || records.length === 0) return null

  const selectedCount = selected.size
  const totalCount = allIds.length

  const renderNode = (node: RecNode, depth: number = 0) => {
    const hasChildren = node.children.length > 0
    const isExpanded = expanded.has(node.id)
    const childIds = getAllIds(node)
    const checkedCount = childIds.filter(id => selected.has(id)).length
    const isChecked = checkedCount === childIds.length
    const isIndeterminate = checkedCount > 0 && !isChecked

    return (
      <Box key={node.id}>
        <Box
          sx={{
            display: 'flex', alignItems: 'center',
            pl: depth * 2,
            py: 0,
            '&:hover': { bgcolor: 'action.hover' },
            borderRadius: 0.5,
          }}
        >
          {hasChildren ? (
            <IconButton size="small" onClick={() => toggleExpand(node.id)} sx={{ p: 0.2 }}>
              {isExpanded
                ? <ExpandMoreOutlined sx={{ fontSize: 15 }} />
                : <ChevronRightOutlined sx={{ fontSize: 15 }} />}
            </IconButton>
          ) : (
            <Box sx={{ width: 22 }} />
          )}
          <Checkbox
            size="small"
            checked={isChecked}
            indeterminate={isIndeterminate}
            onChange={() => toggleNode(node)}
            sx={{ p: 0.2 }}
          />
          <Typography
            variant="body2"
            sx={{
              fontSize: 12,
              color: isChecked || isIndeterminate ? 'text.primary' : 'text.disabled',
              cursor: 'pointer',
              userSelect: 'none',
              lineHeight: 1.3,
            }}
            onClick={() => toggleNode(node)}
          >
            {node.name}
          </Typography>
          {isPeriods && node.level && (
            <Chip
              label={node.level}
              size="small"
              variant="outlined"
              sx={{ ml: 0.5, height: 14, fontSize: 9, '& .MuiChip-label': { px: 0.4 } }}
            />
          )}
        </Box>
        {hasChildren && (
          <Collapse in={isExpanded}>
            {node.children.map(c => renderNode(c, depth + 1))}
          </Collapse>
        )}
      </Box>
    )
  }

  return (
    <Box sx={{ pl: 9, pb: 1 }}>
      <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 0.3 }}>
        {t('sheet.visibleRecords')} ({selectedCount}/{totalCount})
      </Typography>

      {/* Quick buttons */}
      <Box sx={{ display: 'flex', gap: 0.5, mb: 0.5, flexWrap: 'wrap' }}>
        <Chip label={t('sheet.selectAll')} size="small" onClick={selectAll}
          color={selectedCount === totalCount ? 'primary' : 'default'} variant="outlined"
          sx={{ height: 20, fontSize: 11 }} />
        {isPeriods && (
          <>
            <Chip label={t('grid.months')} size="small" onClick={() => selectByLevel('M')}
              variant="outlined" sx={{ height: 20, fontSize: 11 }} />
            <Chip label={t('grid.quarters')} size="small" onClick={() => selectByLevel('Q')}
              variant="outlined" sx={{ height: 20, fontSize: 11 }} />
            <Chip label={t('grid.halfyears')} size="small" onClick={() => selectByLevel('H')}
              variant="outlined" sx={{ height: 20, fontSize: 11 }} />
            <Chip label={t('grid.years')} size="small" onClick={() => selectByLevel('Y')}
              variant="outlined" sx={{ height: 20, fontSize: 11 }} />
          </>
        )}
        <Chip label={t('sheet.selectNone')} size="small" onClick={selectNone}
          variant="outlined" sx={{ height: 20, fontSize: 11 }} />
      </Box>

      {/* Tree */}
      <Box sx={{
        maxHeight: 350, overflow: 'auto',
        border: '1px solid', borderColor: 'divider',
        borderRadius: 1, p: 0.5,
      }}>
        {tree.map(n => renderNode(n, 0))}
      </Box>
    </Box>
  )
}
