# Gantt Chart Template

Formula-driven Excel Gantt chart tracker.

| File | What it is |
|---|---|
| `Project_Execution_Gantt.xlsx` | Completed Gantt for the 1.8M SMC Reflector + 4-Port C-Band Feed System + Metal Structure order (500 sets), start 15-Oct-2026, Mon–Sat working week, 30 linked tasks |
| `Gantt_Chart_Template.xlsx` | Blank-style template with a small sample project |
| `build_gantt_template.py` | Generator: `python3 build_gantt_template.py execution|template out.xlsx` (needs `openpyxl`) |

## Sheets
- **Project Info** – project details, weekly-off setting, schedule health, delay vs target/plan, per-component status, statistics, colour legend.
- **Tasks** – 300 rows ready. Enter task, duration (working days), up to 2 predecessors with link type **FS / SS / FF** and lag. Update Status, Actual Start/End, % Complete, Revised Duration; delays propagate to dependent tasks.
- **Gantt** – Day / Week / Month view (cell B2), dependency markers, weekly-off + holiday shading, status-date column, red overshoot segments, orange delayed-by-predecessor tasks.
- **Holidays** – editable non-working dates.
- **Assumptions** (execution file) – rates, quantities and the basis of every duration.

Open in desktop Excel (2010+); values recalculate on open.
