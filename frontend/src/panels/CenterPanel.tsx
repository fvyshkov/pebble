import { useState, useCallback } from 'react'
import type { TreeSelection } from '../types'
import EmptyState from '../components/EmptyState'
import ModelSettings from '../features/model/ModelSettings'
import AnalyticSettings from '../features/analytic/AnalyticSettings'
import AnalyticFields from '../features/analytic/AnalyticFields'
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
    <div className="panel-center">
      {selection.type === 'model' && (
        <ModelSettings modelId={selection.id} onRefresh={onRefresh} />
      )}
      {selection.type === 'analytic' && (
        <>
          <AnalyticSettings analyticId={selection.id} onRefresh={onInnerRefresh} />
          <AnalyticFields analyticId={selection.id} key={`fields-${innerKey}`} />
          <AnalyticRecordsGrid analyticId={selection.id} key={`records-${innerKey}`} />
        </>
      )}
      {selection.type === 'sheet' && (
        <SheetSettings sheetId={selection.id} modelId={selection.modelId} />
      )}
    </div>
  )
}
