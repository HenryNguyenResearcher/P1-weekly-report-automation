"""
Generate realistic sample CSV data for the Weekly Project Status Report demo.
Produces: tasks.csv, budget.csv, team.csv in the ../data/ directory.
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"

TEAM_MEMBERS = ["Alice Chen", "Bob Martinez", "Carol Kim", "David Nguyen", "Emma Wilson"]
TASK_CATEGORIES = ["Design", "Development", "Testing", "Documentation", "Research", "Deployment"]
STATUSES = ["Done", "In Progress", "In Progress", "In Progress", "Overdue", "Not Started"]
BUDGET_CATEGORIES = ["Design", "Development", "Testing", "Infrastructure", "Marketing", "Operations"]

TASK_TEMPLATES = [
    "Update {} module", "Review {} requirements", "Fix {} bug",
    "Implement {} feature", "Write {} tests", "Deploy {} to staging",
    "Refactor {} code", "Document {} API", "Research {} options",
    "Optimize {} performance", "Configure {} pipeline", "Audit {} logs",
    "Migrate {} data", "Set up {} environment", "Integrate {} service",
    "Validate {} output", "Clean up {} files", "Present {} results",
    "Schedule {} meeting", "Prepare {} report",
]


def random_date(days_back: int = 14, days_ahead: int = 7) -> str:
    today = date.today()
    delta = random.randint(-days_back, days_ahead)
    return (today + timedelta(days=delta)).strftime("%Y-%m-%d")


def generate_tasks(n: int = 20) -> None:
    rows = []
    for i in range(1, n + 1):
        status = random.choice(STATUSES)
        completion = {"Done": 100, "In Progress": random.randint(20, 80),
                      "Not Started": 0, "Overdue": random.randint(5, 50)}[status]
        template = random.choice(TASK_TEMPLATES)
        category = random.choice(TASK_CATEGORIES)
        rows.append({
            "Task ID": f"T-{i:03d}",
            "Task Name": template.format(category),
            "Assigned To": random.choice(TEAM_MEMBERS),
            "Status": status,
            "Due Date": random_date(),
            "Completion %": completion,
        })

    with open(DATA_DIR / "tasks.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"  tasks.csv        — {n} rows")


def generate_budget() -> None:
    rows = []
    for cat in BUDGET_CATEGORIES:
        allocated = random.randint(3000, 15000)
        spent_pct = random.uniform(0.35, 0.95)
        spent = round(allocated * spent_pct, 2)
        rows.append({
            "Category": cat,
            "Allocated ($)": allocated,
            "Spent ($)": spent,
            "Remaining ($)": round(allocated - spent, 2),
        })

    with open(DATA_DIR / "budget.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"  budget.csv       — {len(rows)} categories")


def generate_team(task_file: Path) -> None:
    counts: dict[str, dict] = {
        m: {"Member": m, "Total Tasks": 0, "Completed": 0, "In Progress": 0, "Overdue": 0}
        for m in TEAM_MEMBERS
    }
    with open(task_file, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            member = row["Assigned To"]
            counts[member]["Total Tasks"] += 1
            status = row["Status"]
            if status == "Done":
                counts[member]["Completed"] += 1
            elif status == "In Progress":
                counts[member]["In Progress"] += 1
            elif status == "Overdue":
                counts[member]["Overdue"] += 1

    rows = list(counts.values())
    with open(DATA_DIR / "team.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"  team.csv         — {len(rows)} members")


if __name__ == "__main__":
    DATA_DIR.mkdir(exist_ok=True)
    print("Generating sample data...")
    generate_tasks(20)
    generate_budget()
    generate_team(DATA_DIR / "tasks.csv")
    print("Done. Files saved to:", DATA_DIR.resolve())
