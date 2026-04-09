export interface Model {
  id: string
  name: string
  description: string
  created_at: string
  updated_at: string
}

export interface Analytic {
  id: string
  model_id: string
  name: string
  code: string
  icon: string
  is_periods: number
  data_type: string  // sum | percent | string | quantity
  period_types: string
  period_start: string | null
  period_end: string | null
  sort_order: number
}

export interface AnalyticField {
  id: string
  analytic_id: string
  name: string
  code: string
  data_type: 'string' | 'number' | 'percent' | 'money' | 'date'
  sort_order: number
}

export interface AnalyticRecord {
  id: string
  analytic_id: string
  parent_id: string | null
  sort_order: number
  data_json: string
}

export interface Sheet {
  id: string
  model_id: string
  name: string
  created_at: string
  updated_at: string
}

export interface SheetAnalytic {
  id: string
  sheet_id: string
  analytic_id: string
  sort_order: number
  is_fixed: number
  fixed_record_id: string | null
  analytic_name?: string
  analytic_icon?: string
}

export interface CellData {
  id: string
  sheet_id: string
  coord_key: string
  value: string | null
  data_type: string
  rule: string
  formula: string
}

export interface TreeSelection {
  type: 'model' | 'sheet' | 'analytic'
  id: string
  modelId: string
}
