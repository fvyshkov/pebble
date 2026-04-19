import { createContext, useContext } from 'react'
import { usePendingStore, type PendingOp } from './usePendingStore'
import * as api from '../api'

interface PendingCtx {
  isDirty: boolean
  addOp: (op: Omit<PendingOp, 'ts'>) => void
  getOverrides: (key: string) => Record<string, any> | null
  flush: () => Promise<void>
  clearAll: () => void
}

const Ctx = createContext<PendingCtx>(null!)

export function PendingProvider({ children, onFlushed }: { children: React.ReactNode; onFlushed?: () => void }) {
  const store = usePendingStore()

  const flush = async () => {
    const sorted = store.ops
    for (const op of sorted) {
      try {
        switch (op.type) {
          case 'updateModel':
            await api.updateModel(op.id, op.data as any)
            break
          case 'updateAnalytic':
            await api.updateAnalytic(op.id, op.data as any)
            break
          case 'updateField':
            await api.updateField(op.parentId!, op.id, op.data as any)
            break
          case 'updateRecord':
            await api.updateRecord(op.parentId!, op.id, op.data as any)
            break
          case 'updateSheet':
            await api.updateSheet(op.id, op.data as any)
            break
          case 'putIndicatorRules':
            await api.putIndicatorRules(op.id, op.parentId!, op.data as any)
            break
        }
      } catch (e) {
        console.error('Failed to save op', op, e)
      }
    }
    store.clearAll()
    onFlushed?.()
  }

  return (
    <Ctx.Provider value={{ isDirty: store.isDirty, addOp: store.addOp, getOverrides: store.getOverrides, flush, clearAll: store.clearAll }}>
      {children}
    </Ctx.Provider>
  )
}

export function usePending() {
  return useContext(Ctx)
}
