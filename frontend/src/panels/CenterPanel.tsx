import { useState, useCallback } from 'react'
import type { TreeSelection } from '../types'
import EmptyState from '../components/EmptyState'
import ModelSettings from '../features/model/ModelSettings'
import AnalyticRecordsGrid from '../features/analytic/AnalyticRecordsGrid'
import SheetSettings from '../features/sheet/SheetSettings'

interface Props {
  selection: TreeSelection | null
  onRefresh: () => void
}

export default function CenterPanel({ selection, onRefresh }: Props) {
  const [innerKey, setInnerKey] = useState(0)
  const onInnerRefresh = useCallback(() => {
    setInnerKey(k => k + 1)
    onRefresh()
  }, [onRefresh])

  if (!selection) return <div className="panel-center"><EmptyState /></div>

  return (
    <div className="panel-center" style={selection.type === 'analytic' ? { display: 'flex', flexDirection: 'column', overflow: 'hidden' } : undefined}>
      {selection.type === 'model' && (
        <ModelSettings modelId={selection.id} onRefresh={onRefresh} />
      )}
      {selection.type === 'analytic' && (
        <AnalyticRecordsGrid
          analyticId={selection.id}
          modelId={selection.modelId}
          onRefresh={onInnerRefresh}
          key={`records-${innerKey}`}
        />
      )}
      {selection.type === 'sheet' && (
        <SheetSettings sheetId={selection.id} modelId={selection.modelId} />
      )}
    </div>
  )
}
