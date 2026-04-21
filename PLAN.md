# Feature Plan — Chat & Grid UX Improvements

## P1. Copy button on user prompts in chat [DONE]
Copy button positioned left of blue bubbles, right of white bubbles. Shows on hover.

## P2. Grid context menu → Draw chart from selection [DONE]
- Right-click any cell/range → context menu with bar/line/pie chart options
- Row labels become series, column headers become categories
- Chart renders in floating overlay on grid (closeable with X)
- Chart.js loaded dynamically from CDN

## P3. Cell change history in context menu [ALREADY EXISTED]
- cell_history table was already in the DB
- Context menu "История изменений" option was already implemented
- Shows who, when, old value, new value

## P4. Chat stop button [DONE]
- Send button becomes red stop icon while loading
- AbortController cancels the SSE fetch

## P5. Analytic value drag — full cell area [ALREADY WORKS]
- The row header `<td>` is already `draggable` with `cursor: grab`
- Entire cell area is draggable, not just the text

## P6. File/image attachments in chat [DONE]
- Paste images from clipboard or drop files onto chat
- Attached files show as chips with icons above input (removable)
- Images are base64-encoded and sent as Claude vision content
- Excel files trigger import dialog flow

## Tests [DONE]
- test_chat_stop_button
- test_chat_copy_button  
- test_chat_prompt_history
- test_grid_context_menu_chart_options
- All 10 chat tests pass (111s)
- Full test suite running...
