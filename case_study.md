# Case Study: Automating the Friday Report

**Tool**: Python + openpyxl  
**Time saved**: ~90 minutes → 0.12 seconds per report  
**Skill demonstrated**: Automation VA / Data Operations

---

## The Problem

Every Friday afternoon, a project manager at a small digital agency did the same manual task:
open three spreadsheets, copy numbers into a report template, update charts, fix formatting, then send.

With 20+ active tasks and five team members, this took **90 minutes** — time that could have been spent on client work. Worse, copy-paste errors occasionally made it into the final report.

This is exactly the kind of repetitive, high-effort task that a skilled VA can eliminate permanently.

---

## The Solution

I built a Python automation script that:

1. **Reads** three source CSV files (tasks, budget, team workload)
2. **Calculates** KPIs automatically (completion %, budget used %, overdue count)
3. **Generates** a formatted four-sheet Excel report with charts and conditional formatting
4. **Saves** the file with today's date in the filename — ready to send

The entire process runs in **one command**:

```bash
python src/report_generator.py
```

---

## What the Report Contains

- **Summary sheet**: Four KPI cards at a glance (tasks done, budget used, overdue, team utilisation)
- **Tasks sheet**: Colour-coded table — green for done, yellow for in-progress, red for overdue
- **Budget sheet**: Table + clustered bar chart comparing allocated vs. spent per category
- **Team sheet**: Workload table + stacked bar chart showing each member's task distribution

---

## Results

| Metric | Manual Process | Automated |
| --- | --- | --- |
| Time per report | ~90 min | 0.12 sec |
| Copy-paste errors | Occasional | Zero |
| Rows handled | ~20 | 500+ |
| Consistency | Varied | Always identical |

The 90-minute weekly task became a one-second command.

---

## Why This Matters for a VA Client

As a Virtual Assistant, I don't just complete tasks — I look for ways to make them disappear.

This type of automation (reading raw data → producing clean output) applies to dozens of common VA workflows:

- Weekly/monthly business reports
- CRM data exports → formatted summaries
- Invoice or expense tracking dashboards
- Client onboarding checklists

If you're spending more than 30 minutes a week on a report that follows the same structure every time, it can almost certainly be automated.

---

*Full source code: [github.com/HenryNguyenResearcher/P1-weekly-report-automation](https://github.com/HenryNguyenResearcher/P1-weekly-report-automation)*  
*Contact: [hoangnvm.hust@gmail.com](mailto:hoangnvm.hust@gmail.com)*
