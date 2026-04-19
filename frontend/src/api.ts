import type { Model, Analytic, AnalyticField, AnalyticRecord, Sheet, SheetAnalytic, CellData } from './types'

const BASE = '/api'

async function json<T>(r: Response): Promise<T> {
  if (!r.ok) throw new Error(`${r.status} ${r.statusText}`)
  return r.json()
}

// Models
export const listModels = () => fetch(`${BASE}/models`).then(r => json<Model[]>(r))
export const createModel = (name: string) =>
  fetch(`${BASE}/models`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name }) }).then(r => json<Model>(r))
export const updateModel = (id: string, data: { name: string; description: string }) =>
  fetch(`${BASE}/models/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<Model>(r))
export const deleteModel = (id: string) =>
  fetch(`${BASE}/models/${id}`, { method: 'DELETE' }).then(r => json<any>(r))
export const getModelTree = (id: string) =>
  fetch(`${BASE}/models/${id}/tree`).then(r => json<any>(r))

// Analytics
export const listAnalytics = (modelId: string) =>
  fetch(`${BASE}/analytics/by-model/${modelId}`).then(r => json<Analytic[]>(r))
export const createAnalytic = (data: Partial<Analytic>) =>
  fetch(`${BASE}/analytics`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<Analytic>(r))
export const getAnalytic = (id: string) =>
  fetch(`${BASE}/analytics/${id}`).then(r => json<Analytic>(r))
export const updateAnalytic = (id: string, data: any) =>
  fetch(`${BASE}/analytics/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<Analytic>(r))
export const deleteAnalytic = (id: string) =>
  fetch(`${BASE}/analytics/${id}`, { method: 'DELETE' }).then(r => json<any>(r))

// Fields
export const listFields = (analyticId: string) =>
  fetch(`${BASE}/analytics/${analyticId}/fields`).then(r => json<AnalyticField[]>(r))
export const createField = (analyticId: string, data: Partial<AnalyticField>) =>
  fetch(`${BASE}/analytics/${analyticId}/fields`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<AnalyticField>(r))
export const updateField = (analyticId: string, fieldId: string, data: Partial<AnalyticField>) =>
  fetch(`${BASE}/analytics/${analyticId}/fields/${fieldId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<AnalyticField>(r))
export const deleteField = (analyticId: string, fieldId: string) =>
  fetch(`${BASE}/analytics/${analyticId}/fields/${fieldId}`, { method: 'DELETE' }).then(r => json<any>(r))

// Records
export const listRecords = (analyticId: string) =>
  fetch(`${BASE}/analytics/${analyticId}/records`).then(r => json<AnalyticRecord[]>(r))
export const createRecord = (analyticId: string, data: any) =>
  fetch(`${BASE}/analytics/${analyticId}/records`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<AnalyticRecord>(r))
export const updateRecord = (analyticId: string, recordId: string, data: any) =>
  fetch(`${BASE}/analytics/${analyticId}/records/${recordId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<AnalyticRecord>(r))
export const deleteRecord = (analyticId: string, recordId: string) =>
  fetch(`${BASE}/analytics/${analyticId}/records/${recordId}`, { method: 'DELETE' }).then(r => json<any>(r))
export const generatePeriods = (analyticId: string) =>
  fetch(`${BASE}/analytics/${analyticId}/generate-periods`, { method: 'POST' }).then(r => json<AnalyticRecord[]>(r))

// Sheets
export const listSheets = (modelId: string) =>
  fetch(`${BASE}/sheets/by-model/${modelId}`).then(r => json<Sheet[]>(r))
export const createSheet = (data: { model_id: string; name: string }) =>
  fetch(`${BASE}/sheets`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<Sheet>(r))
export const updateSheet = (id: string, data: { name: string }) =>
  fetch(`${BASE}/sheets/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<Sheet>(r))
export const deleteSheet = (id: string) =>
  fetch(`${BASE}/sheets/${id}`, { method: 'DELETE' }).then(r => json<any>(r))
export const reorderSheets = (modelId: string, orderedIds: string[]) =>
  fetch(`${BASE}/sheets/reorder/${modelId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ordered_ids: orderedIds }) }).then(r => json<any>(r))

// Sheet Analytics
export const listSheetAnalytics = (sheetId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics`).then(r => json<SheetAnalytic[]>(r))
export const addSheetAnalytic = (sheetId: string, data: { analytic_id: string; sort_order: number }) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<SheetAnalytic>(r))
export const removeSheetAnalytic = (sheetId: string, saId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics/${saId}`, { method: 'DELETE' }).then(r => json<any>(r))
export const reorderSheetAnalytics = (sheetId: string, orderedIds: string[]) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics-reorder`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ordered_ids: orderedIds }) }).then(r => json<any>(r))

// Main analytic + indicator formula rules
export const getMainAnalytic = (sheetId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/main-analytic`).then(r => json<{ analytic_id: string | null }>(r))
export const setMainAnalytic = (sheetId: string, analyticId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/main-analytic`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ analytic_id: analyticId }) }).then(r => json<any>(r))

export interface ScopedRule { id?: string; scope: Record<string, string>; priority: number; formula: string }
export interface IndicatorRules { leaf: string; consolidation: string; scoped: ScopedRule[] }
export const getIndicatorRules = (sheetId: string, indicatorId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/indicators/${indicatorId}/rules`).then(r => json<IndicatorRules>(r))
export const putIndicatorRules = (sheetId: string, indicatorId: string, data: IndicatorRules) =>
  fetch(`${BASE}/sheets/${sheetId}/indicators/${indicatorId}/rules`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<any>(r))
export const promoteCellToRule = (sheetId: string, indicatorId: string, coordKey: string, formula: string, priority = 100) =>
  fetch(`${BASE}/sheets/${sheetId}/indicators/${indicatorId}/rules/promote-cell`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ coord_key: coordKey, formula, priority }) }).then(r => json<any>(r))
export const getResolvedFormulas = (sheetId: string, coordKeys: string[]) =>
  fetch(`${BASE}/sheets/${sheetId}/cells/resolved-formulas`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ coord_keys: coordKeys }) }).then(r => json<{ coord_key: string; formula: string; source: string; kind: string }[]>(r))

// View settings
export const getViewSettings = (sheetId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/view-settings`).then(r => json<any>(r))
export const saveViewSettings = (sheetId: string, settings: any) =>
  fetch(`${BASE}/sheets/${sheetId}/view-settings`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ settings }) }).then(r => json<any>(r))

// Cells
export const getCells = (sheetId: string, userId?: string) =>
  fetch(`${BASE}/cells/by-sheet/${sheetId}${userId ? `?user_id=${userId}` : ''}`).then(r => json<CellData[]>(r))
export const saveCells = (sheetId: string, cells: { coord_key: string; value?: string | null; data_type?: string; user_id?: string; rule?: string; formula?: string }[], noRecalc = false) =>
  fetch(`${BASE}/cells/by-sheet/${sheetId}${noRecalc ? '?no_recalc=true' : ''}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cells }) }).then(r => json<any>(r))
export const calculateSheet = (sheetId: string) =>
  fetch(`${BASE}/cells/calculate/${sheetId}`, { method: 'POST' }).then(r => json<any>(r))
export const getCellsPartial = (sheetId: string, coordKeys: string[], userId?: string) =>
  fetch(`${BASE}/cells/by-sheet/${sheetId}/partial${userId ? `?user_id=${userId}` : ''}`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ coord_keys: coordKeys }),
  }).then(r => json<CellData[]>(r))
export const calculateModelStream = (
  modelId: string,
  onProgress: (data: { phase: string; sheet?: string; done?: number; total_sheets?: number; computed?: number; total_cells?: number }) => void,
): Promise<void> =>
  fetch(`${BASE}/cells/calculate-model/${modelId}/stream`, { method: 'POST' })
    .then(async res => {
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
      const reader = res.body!.getReader()
      const decoder = new TextDecoder()
      let buffer = ''
      while (true) {
        const { done, value } = await reader.read()
        if (done) break
        buffer += decoder.decode(value, { stream: true })
        const lines = buffer.split('\n')
        buffer = lines.pop() || ''
        for (const line of lines) {
          if (line.startsWith('data: ')) {
            try { onProgress(JSON.parse(line.slice(6))) } catch {}
          }
        }
      }
    })
export const getCellHistory = (sheetId: string, coordKey: string) =>
  fetch(`${BASE}/cells/history/${sheetId}/${encodeURIComponent(coordKey)}`).then(r => json<any[]>(r))
export const getModelHistory = (modelId: string, limit = 10) =>
  fetch(`${BASE}/cells/model-history/${modelId}?limit=${limit}`).then(r => json<any[]>(r))
export const undoChanges = (modelId: string, historyId: string) =>
  fetch(`${BASE}/cells/undo/${modelId}`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ history_id: historyId }) }).then(r => json<any>(r))
export const clearHistory = (modelId: string) =>
  fetch(`${BASE}/cells/model-history/${modelId}`, { method: 'DELETE' }).then(r => json<any>(r))

// Analytic record permissions
export const getAnalyticPermissions = (userId: string) =>
  fetch(`${BASE}/users/${userId}/analytic-permissions`).then(r => json<any[]>(r))
export const getAllowedRecords = (userId: string, sheetId: string) =>
  fetch(`${BASE}/users/${userId}/allowed-records/${sheetId}`).then(r => json<Record<string, string[]>>(r))
export const setAnalyticPermission = (data: { user_id: string; analytic_id: string; record_id: string; can_view: boolean; can_edit: boolean }) =>
  fetch(`${BASE}/users/analytic-permissions/set`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<any>(r))

// Export
export const exportModel = (modelId: string) =>
  `${BASE}/excel/models/${modelId}/export`

// Users
export const listUsers = () => fetch(`${BASE}/users`).then(r => json<any[]>(r))
export const createUser = (username: string) =>
  fetch(`${BASE}/users`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username }) }).then(r => json<any>(r))
export const updateUser = (id: string, username: string) =>
  fetch(`${BASE}/users/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username }) }).then(r => json<any>(r))
export const deleteUser = (id: string) =>
  fetch(`${BASE}/users/${id}`, { method: 'DELETE' }).then(r => json<any>(r))
export const resetPassword = (id: string, password: string) =>
  fetch(`${BASE}/users/${id}/reset-password`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ password }) }).then(r => json<any>(r))
export const setAdmin = (id: string, canAdmin: boolean) =>
  fetch(`${BASE}/users/${id}/admin`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ can_admin: canAdmin }) }).then(r => json<any>(r))
export const getSheetPermissions = (sheetId: string) =>
  fetch(`${BASE}/users/permissions/by-sheet/${sheetId}`).then(r => json<any[]>(r))
export const setSheetPermission = (sheetId: string, data: { user_id: string; can_view: boolean; can_edit: boolean }) =>
  fetch(`${BASE}/users/permissions/by-sheet/${sheetId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<any>(r))
export const getAccessibleSheets = (userId: string) =>
  fetch(`${BASE}/users/${userId}/accessible-sheets`).then(r => json<any[]>(r))
export const getAllPermissions = (userId: string) =>
  fetch(`${BASE}/users/${userId}/all-permissions`).then(r => json<any[]>(r))

// Import model from Excel
export const importExcelModel = (file: File, modelName: string) => {
  const fd = new FormData()
  fd.append('file', file)
  fd.append('model_name', modelName)
  return fetch(`${BASE}/import/excel`, { method: 'POST', body: fd }).then(r => json<any>(r))
}

// Import model from Excel with streaming progress
export const importExcelModelStream = (
  file: File, modelName: string,
  onMessage: (msg: string, data?: any) => void,
): Promise<any> => {
  const fd = new FormData()
  fd.append('file', file)
  fd.append('model_name', modelName)
  return new Promise((resolve, reject) => {
    fetch(`${BASE}/import/excel-stream`, { method: 'POST', body: fd })
      .then(async res => {
        if (!res.ok) throw new Error(`${res.status} ${res.statusText}`)
        const reader = res.body!.getReader()
        const decoder = new TextDecoder()
        let buffer = ''
        let lastData: any = null
        while (true) {
          const { done, value } = await reader.read()
          if (done) break
          buffer += decoder.decode(value, { stream: true })
          const lines = buffer.split('\n')
          buffer = lines.pop() || ''
          for (const line of lines) {
            if (line.startsWith('data: ')) {
              try {
                const parsed = JSON.parse(line.slice(6))
                onMessage(parsed.message, parsed)
                if (parsed.done) lastData = parsed
              } catch {}
            }
          }
        }
        resolve(lastData)
      })
      .catch(reject)
  })
}

// Sheet data Excel export/import
export const exportSheetExcelUrl = (sheetId: string) => `${BASE}/excel/sheets/${sheetId}/export`
export const importSheetExcel = (sheetId: string, file: File) => {
  const fd = new FormData()
  fd.append('file', file)
  return fetch(`${BASE}/excel/sheets/${sheetId}/import`, { method: 'PUT', body: fd }).then(r => json<any>(r))
}

// Excel
export const exportExcelUrl = (analyticId: string) => `${BASE}/excel/analytics/${analyticId}/export`
export const importExcel = (analyticId: string, file: File) => {
  const fd = new FormData()
  fd.append('file', file)
  return fetch(`${BASE}/excel/analytics/${analyticId}/import`, { method: 'POST', body: fd }).then(r => json<AnalyticRecord[]>(r))
}

// Chat
export interface ChatAction {
  type: string
  [k: string]: any
}
export const chatMessage = (
  messages: { role: string; content: any }[],
  context: { current_model_id?: string | null; current_sheet_id?: string | null; user_id?: string | null },
) =>
  fetch(`${BASE}/chat/message`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ messages, context }),
  }).then(r => json<{ message: string; actions: ChatAction[] }>(r))
