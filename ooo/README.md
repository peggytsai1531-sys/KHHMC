# OOO 歷史報表

每日 T114 Out of Order 報表的分析結果，依月份存放於 `history/YYYY-MM/`。

- `YYYY-MM-DD_rooms.csv`：當日已生效 OOO 房號明細（含連續天數）
- `YYYY-MM-DD_diff.csv`：與前一份報表的差異（新增／解除／持續）
- `YYYY-MM-DD.json`：當日所有房號（含未來預排，active=false）
- `summary.csv`：當月每日摘要
- `OOO_YYYY-MM.xlsx`：當月 Excel（每日摘要、前日差異、每日明細各一張工作表）

產生方式：`python3 -I ooo/build_history.py <ooo_result.json>`（輸入為 classify_ooo.py 的 JSON）。
