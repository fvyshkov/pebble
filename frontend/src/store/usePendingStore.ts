import { useState, useCallback, useEffect, useMemo } from 'react'

const STORAGE_KEY = 'pebble_pending'

export interface PendingOp {
  key: string // unique key: "model:id", "analytic:id", "field:id", etc.
  type: 'updateModel' | 'updateAnalytic' | 'updateField' | 'updateSheet' | 'updateRecord'
  id: string
  parentId?: string // e.g. analyticId for fields/records
  data: Record<string, any>
  ts: number
}

function loadPending(): Record<string, PendingOp> {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : {}
  } catch { return {} }
}

function savePending(ops: Record<string, PendingOp>) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(ops))
}

export function usePendingStore() {
  const [ops, setOps] = useState<Record<string, PendingOp>>(loadPending)

  useEffect(() => { savePending(ops) }, [ops])

  const isDirty = useMemo(() => Object.keys(ops).length > 0, [ops])

  const addOp = useCallback((op: Omit<PendingOp, 'ts'>) => {
    setOps(prev => {
      const existing = prev[op.key]
      const merged: PendingOp = existing
        ? { ...existing, data: { ...existing.data, ...op.data }, ts: Date.now() }
        : { ...op, ts: Date.now() }
      return { ...prev, [op.key]: merged }
    })
  }, [])

  const getOverrides = useCallback((key: string): Record<string, any> | null => {
    return ops[key]?.data ?? null
  }, [ops])

  const clearAll = useCallback(() => {
    setOps({})
    localStorage.removeItem(STORAGE_KEY)
  }, [])

  const allOps = useMemo(() => Object.values(ops).sort((a, b) => a.ts - b.ts), [ops])

  return { isDirty, ops: allOps, addOp, getOverrides, clearAll }
}
