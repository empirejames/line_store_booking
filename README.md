# LINE Store Booking & Bookkeeping System

專屬 LINE 官方帳號記帳與營收分析系統。

## 系統架構
- **前端報表與圖表**：託管於 Vercel / GitHub Pages (`chart.html`, `dividend.html`, `time_chart.html`, `report.html`)
- **後端 API 與試算表串接**：Google Apps Script (`backend.gs`)
- **定時推播排程**：GitHub Actions
  - `LINE Daily Reminder`：每日 10:30、16:00 發送填表提醒
  - `LINE Daily Summary`：每日 21:00 發送當日業績戰報 Flex Message

## 維護備忘
- GitHub Actions 排程若連續 60 天無 commit，將會進入 `disabled_inactivity` 狀態。定期維護或提交 commit 可保持排程活躍。
