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

// Sheet Analytics
export const listSheetAnalytics = (sheetId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics`).then(r => json<SheetAnalytic[]>(r))
export const addSheetAnalytic = (sheetId: string, data: { analytic_id: string; sort_order: number }) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(data) }).then(r => json<SheetAnalytic>(r))
export const removeSheetAnalytic = (sheetId: string, saId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics/${saId}`, { method: 'DELETE' }).then(r => json<any>(r))
export const reorderSheetAnalytics = (sheetId: string, orderedIds: string[]) =>
  fetch(`${BASE}/sheets/${sheetId}/analytics-reorder`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ordered_ids: orderedIds }) }).then(r => json<any>(r))

// View settings
export const getViewSettings = (sheetId: string) =>
  fetch(`${BASE}/sheets/${sheetId}/view-settings`).then(r => json<any>(r))
export const saveViewSettings = (sheetId: string, settings: any) =>
  fetch(`${BASE}/sheets/${sheetId}/view-settings`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ settings }) }).then(r => json<any>(r))

// Cells
export const getCells = (sheetId: string) =>
  fetch(`${BASE}/cells/by-sheet/${sheetId}`).then(r => json<CellData[]>(r))
export const saveCells = (sheetId: string, cells: { coord_key: string; value?: string | null; data_type?: string; user_id?: string; rule?: string; formula?: string }[]) =>
  fetch(`${BASE}/cells/by-sheet/${sheetId}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ cells }) }).then(r => json<any>(r))
export const getCellHistory = (sheetId: string, coordKey: string) =>
  fetch(`${BASE}/cells/history/${sheetId}/${encodeURIComponent(coordKey)}`).then(r => json<any[]>(r))

// Users
export const listUsers = () => fetch(`${BASE}/users`).then(r => json<any[]>(r))
export const createUser = (username: string) =>
  fetch(`${BASE}/users`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username }) }).then(r => json<any>(r))
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

// Import model from Excel
export const importExcelModel = (file: File, modelName: string) => {
  const fd = new FormData()
  fd.append('file', file)
  fd.append('model_name', modelName)
  return fetch(`${BASE}/import/excel`, { method: 'POST', body: fd }).then(r => json<any>(r))
}

// Excel
export const exportExcelUrl = (analyticId: string) => `${BASE}/excel/analytics/${analyticId}/export`
export const importExcel = (analyticId: string, file: File) => {
  const fd = new FormData()
  fd.append('file', file)
  return fetch(`${BASE}/excel/analytics/${analyticId}/import`, { method: 'POST', body: fd }).then(r => json<AnalyticRecord[]>(r))
}
