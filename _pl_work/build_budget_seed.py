from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime
from pathlib import Path


BASE = Path(__file__).resolve().parent
LEDGER = BASE / "pl_ledger.json"
OUT = BASE / "budget_seed.json"


def month_label(period: str) -> str:
    year, month = period.split("-")
    return f"{month}/{year}"


def next_periods_from(last_period: str, through_month: int = 12) -> list[str]:
    year, month = [int(part) for part in last_period.split("-")]
    return [f"{year}-{m:02d}" for m in range(month + 1, through_month + 1)]


def main():
    ledger = json.loads(LEDGER.read_text(encoding="utf-8"))
    periods = sorted({row.get("period") for row in ledger if row.get("period")})
    current_year = max(period[:4] for period in periods)
    year_periods = [period for period in periods if period.startswith(current_year)]
    base_periods = year_periods[-4:] if len(year_periods) >= 4 else periods[-4:]
    project_periods = next_periods_from(base_periods[-1])

    dimensions = (
        "client",
        "project",
        "hub",
        "expt",
        "vehicleType",
        "fleetType",
        "account",
        "category",
        "costType",
    )

    grouped = defaultdict(lambda: defaultdict(float))
    for row in ledger:
        period = row.get("period")
        if period not in base_periods:
            continue
        key = tuple(row.get(name) or "" for name in dimensions)
        grouped[key][period] += row.get("value") or 0.0

    base_rows = []
    budget_rows = []
    history_rows = []

    for key, values in sorted(grouped.items()):
        vals = [values.get(period, 0.0) for period in base_periods]
        average = sum(vals) / len(base_periods) if base_periods else 0.0
        if abs(average) <= 0.0001 and not any(abs(v) > 0.0001 for v in vals):
            continue
        record = dict(zip(dimensions, key))
        base_rows.append(
            {
                **record,
                **{period: values.get(period, 0.0) for period in base_periods},
                "average": average,
                "activeMonths": sum(1 for value in vals if abs(value) > 0.0001),
            }
        )
        for period in project_periods:
            budget_rows.append(
                {
                    "scenario": "Budget Média 4M",
                    "year": int(period[:4]),
                    "month": int(period[5:]),
                    "period": period,
                    "monthLabel": month_label(period),
                    **record,
                    "budgetValue": average,
                    "note": "",
                }
            )

    history_grouped = defaultdict(float)
    for row in ledger:
        period = row.get("period")
        if period not in base_periods:
            continue
        key = tuple([period] + [row.get(name) or "" for name in dimensions])
        history_grouped[key] += row.get("value") or 0.0

    for key, value in sorted(history_grouped.items()):
        period = key[0]
        record = dict(zip(dimensions, key[1:]))
        history_rows.append(
            {
                "period": period,
                "monthLabel": month_label(period),
                **record,
                "actualValue": value,
            }
        )

    data = {
        "meta": {
            "generatedAt": datetime.now().isoformat(timespec="seconds"),
            "source": str(LEDGER),
            "basePeriods": base_periods,
            "projectPeriods": project_periods,
            "basePeriodLabels": [month_label(period) for period in base_periods],
            "projectPeriodLabels": [month_label(period) for period in project_periods],
            "baseRows": len(base_rows),
            "budgetRows": len(budget_rows),
            "historyRows": len(history_rows),
        },
        "baseRows": base_rows,
        "budgetRows": budget_rows,
        "historyRows": history_rows,
    }
    OUT.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print(OUT)
    print(json.dumps(data["meta"], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
