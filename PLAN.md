# Pebble — Мини-Anaplan

## Стек
- **Backend**: FastAPI + SQLite (aiosqlite) + openpyxl
- **Frontend**: React 18 + TypeScript + Vite + MUI 5

## Итерация 1
1. Модели (CRUD)
2. Аналитики (CRUD, иерархия, периоды, поля, записи, TreeGrid)
3. Листы (CRUD, привязка аналитик, порядок)
4. Дерево слева (модель > листы/аналитики)
5. Центральная панель — настройки выбранного элемента
6. Excel импорт/экспорт для записей аналитик

## Итерация 2
- Просмотр/ввод данных на листе (pivot-таблица)
- Фиксация аналитик
- Суммирование по иерархии

## Итерация 3
- Доработки UI, правая панель
- Полировка

## Структура проекта
```
pebble/
├── backend/
│   ├── main.py
│   ├── db.py
│   ├── config.py
│   ├── transliterate.py
│   ├── requirements.txt
│   └── routers/
│       ├── models.py
│       ├── analytics.py
│       ├── sheets.py
│       ├── cells.py
│       └── excel_io.py
└── frontend/
    ├── package.json
    ├── vite.config.ts
    ├── tsconfig.json
    ├── index.html
    └── src/
        ├── main.tsx
        ├── App.tsx
        ├── App.css
        ├── api.ts
        ├── types.ts
        ├── utils/
        │   └── transliterate.ts
        ├── components/
        │   ├── Splitter.tsx
        │   ├── IconPickerDialog.tsx
        │   ├── ConfirmDialog.tsx
        │   └── EmptyState.tsx
        ├── panels/
        │   ├── LeftPanel.tsx
        │   ├── CenterPanel.tsx
        │   └── RightPanel.tsx
        └── features/
            ├── model/ModelSettings.tsx
            ├── analytic/
            │   ├── AnalyticSettings.tsx
            │   ├── AnalyticFields.tsx
            │   └── AnalyticRecordsGrid.tsx
            └── sheet/
                └── SheetSettings.tsx
```

## БД (SQLite)
- models, analytics, analytic_fields, analytic_records (parent_id для иерархии)
- sheets, sheet_analytics, cell_data (coord_key = record IDs через "|")
