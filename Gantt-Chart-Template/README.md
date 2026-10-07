# Gantt Chart Template

A formula-driven Excel Gantt chart template (`Gantt_Chart_Template.xlsx`).

## Sheets
- **Project Info** – project name, client, start / target dates, report date; schedule health, delay vs target/plan, task statistics, colour legend.
- **Tasks** – unlimited-style task entry (300 rows ready; copy the last row down for more): duration, up to 2 predecessors, status, actual dates, % complete, revised duration. Calculates planned/forecast dates, duration variance and "delayed by" impact.
- **Gantt** – Day / Week / Month view (cell B2), dependency markers (▶ ↳ ◀ ➜ ◆), weekend/holiday shading, status-date column, red overshoot segments, orange delayed-by-predecessor tasks.
- **Holidays** – editable list of non-working dates (Sat/Sun automatic).

## Rebuild
`python3 build_gantt_template.py [output.xlsx]` (requires `openpyxl`). The file recalculates when opened in Excel.
