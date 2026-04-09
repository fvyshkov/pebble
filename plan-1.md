# Plan: Excel Model Import

## Analysis of "AA models.xlsx"

### Structure
- **7 sheets**: `0` (parameters), `BaaS.1` (lending), `BaaS.2` (deposits), `BaaS.3` (transactions), `BS` (balance sheet), `PL` (P&L), `OPEX+CAPEX` (expenses)
- **Periods**: monthly, Jan 2026 – Dec 2028 (columns starting from D or E)
- **Row dimensions**: indicator names in col A/B, some sheets have product sub-groups (e.g. "Потребительский кредит", "Рассрочка (BNPL)")
- **Cell types**: input cells (theme=7, beige bg) vs formula cells (theme=0, white bg)
- **Formulas**: intra-sheet (`=D13*D14`), cross-sheet (`='0'!D10`, `=-'OPEX+CAPEX'!E19`), aggregation (`=SUM(D28:D31)`), prev-period refs (`=D20+E17-E19`)
- **BS/PL**: 100% formula sheets — purely derived from BaaS.1-3 and OPEX+CAPEX
- **OPEX+CAPEX**: has 3 row dimensions — Product (col A), Expense Type (col B), Line Item (col C)

### Mapping to Pebble model

**Analytics (per sheet):**
1. `Периоды` — shared across all sheets, auto-generated monthly periods
2. `Показатели` — each sheet gets its own hierarchical indicator analytic (rows)
3. `OPEX+CAPEX` gets additional analytics: Продукт, Вид расхода, Статья

**Sheets** = Excel tabs → Pebble sheets, each bound to Периоды + its row analytics

**Data**: input cells (beige/theme=7) → rule=manual. Formula cells → stored with computed values.

## Implementation

### Step 1: Backend — import endpoint
- [x] `POST /api/import/excel` — accepts XLSX + model name
- [x] Auto-detect period columns from date headers
- [x] Create model, shared period analytic
- [x] For each Excel sheet: detect row structure, create indicator analytic(s) with hierarchy
- [x] Read cell values (data_only=True), detect input vs formula via theme color
- [x] Store input cells as manual, formula cells with computed values
- [x] If model name already exists, append datetime suffix
- Tested: 7 sheets, 39 periods, 662 records, 21877 cells (5784 manual + 16093 formula)

### Step 2: Frontend — import button + dialog
- [x] Import button in toolbar (upload icon)
- [x] Dialog: file picker + model name input
- [x] Call import endpoint, refresh tree, select new model

### Step 3: Future — formula support improvements
- [ ] Cross-sheet references
- [ ] "Last value" aggregation type for quarters/years
