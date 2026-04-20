# Weekly Project Status Report — Automation Tool

**Python script that turns 3 raw CSV files into a fully formatted Excel report in under 1 second.**

---

## The Problem

A project manager at a small agency spent **90 minutes every Friday** doing the same thing:

1. Open three spreadsheets (tasks, budget, team workload)
2. Copy-paste numbers into a report template by hand
3. Update charts manually
4. Format everything before sending to stakeholders

With 20+ tasks and 5+ team members, this was tedious, error-prone, and a waste of billable time.

---

## The Solution

A Python script that reads the three source files and automatically generates a clean, formatted Excel report — complete with KPI cards, conditional formatting, and charts.

```
python src/report_generator.py
```

**Input** (3 CSV files in `data/`):

| File | Contents |
|------|----------|
| `tasks.csv` | Task ID, Name, Assignee, Status, Due Date, Completion % |
| `budget.csv` | Category, Allocated ($), Spent ($), Remaining ($) |
| `team.csv` | Member, Total Tasks, Completed, In Progress, Overdue |

**Output** (`output/weekly_report_YYYY-MM-DD.xlsx`):

| Sheet | Contents |
|-------|----------|
| Summary | 4 KPI cards + status breakdown table |
| Tasks | Colour-coded task table (green/yellow/red by status) |
| Budget | Formatted table + Allocated vs Spent bar chart |
| Team | Workload table + stacked bar chart |

---

## Results

| Metric | Before (Manual) | After (Automated) |
|--------|----------------|-------------------|
| Time to generate report | ~90 minutes | **0.12 seconds** |
| Data sources consolidated | 3 (manually) | 3 (automatically) |
| Human errors | Common (copy-paste) | **Zero** |
| Rows processed | 20–500+ | **Unlimited** |

> "What used to take the whole Friday morning now runs while I pour my coffee."

---

## Tools Used

- **Python 3.10+**
- **pandas** — data loading and KPI calculation
- **openpyxl** — Excel file creation, formatting, and chart generation

---

## How to Run

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Generate sample data (optional — skip if you have your own CSVs)
python src/generate_sample_data.py

# 3. Generate the report
python src/report_generator.py
```

The output file will appear in the `output/` folder.

---

## Customisation

To use your own data, replace the CSV files in `data/` — the script will automatically adapt to the number of tasks, categories, and team members.

---

*Built as a portfolio project to demonstrate automation skills for Virtual Assistant / Operations roles.*  
*Author: Nguyen Van Minh Hoang | [GitHub](https://github.com/HenryNguyenResearcher) )*
