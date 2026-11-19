"""Read a CSV of support tickets, summarize by category, write an Excel
workbook and a Word report. No network, no secrets; input in, files out."""
from __future__ import annotations

import csv
import sys
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from openpyxl import Workbook


@dataclass
class CategoryStats:
    count: int = 0
    resolved: int = 0
    hours: list[float] = field(default_factory=list)

    @property
    def resolution_rate(self) -> float:
        return self.resolved / self.count if self.count else 0.0

    @property
    def avg_hours_open(self) -> float:
        return sum(self.hours) / len(self.hours) if self.hours else 0.0


def load_tickets(csv_path: Path) -> dict[str, CategoryStats]:
    stats: dict[str, CategoryStats] = {}
    with csv_path.open(newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            cat = row["category"].strip().lower()
            s = stats.setdefault(cat, CategoryStats())
            s.count += 1
            s.resolved += row["resolved"].strip().lower() == "yes"
            s.hours.append(float(row["hours_open"]))
    return stats


def write_excel(stats: dict[str, CategoryStats], out_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "by category"
    ws.append(["category", "tickets", "resolved", "resolution_rate", "avg_hours_open"])
    for cat, s in sorted(stats.items()):
        ws.append([cat, s.count, s.resolved, round(s.resolution_rate, 3), round(s.avg_hours_open, 2)])
    wb.save(out_path)


def write_docx(stats: dict[str, CategoryStats], out_path: Path) -> None:
    doc = Document()
    doc.add_heading("Ticket summary", level=1)
    total = sum(s.count for s in stats.values())
    doc.add_paragraph(f"{total} tickets across {len(stats)} categories.")
    worst = max(stats.items(), key=lambda kv: kv[1].avg_hours_open, default=None)
    for cat, s in sorted(stats.items()):
        doc.add_paragraph(
            f"{cat}: {s.count} tickets, {s.resolution_rate:.0%} resolved, "
            f"average {s.avg_hours_open:.1f} hours open.",
            style="List Bullet",
        )
    if worst:
        doc.add_paragraph(
            f"Slowest category is {worst[0]} at {worst[1].avg_hours_open:.1f} "
            "average hours open. That is the queue to staff first."
        )
    doc.save(out_path)


def main(argv: list[str]) -> int:
    src = Path(argv[1]) if len(argv) > 1 else Path("data/sample-tickets.csv")
    out_dir = Path(argv[2]) if len(argv) > 2 else Path("out")
    out_dir.mkdir(parents=True, exist_ok=True)
    stats = load_tickets(src)
    xlsx, docx = out_dir / "report.xlsx", out_dir / "report.docx"
    write_excel(stats, xlsx)
    write_docx(stats, docx)
    print(f"wrote {xlsx} and {docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
