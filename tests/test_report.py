from pathlib import Path

from openpyxl import load_workbook

from src.report import load_tickets, main, write_excel

SAMPLE = Path(__file__).parent.parent / "data" / "sample-tickets.csv"


def test_load_tickets_aggregates_by_category():
    stats = load_tickets(SAMPLE)
    assert stats["network"].count == 4
    assert stats["network"].resolved == 4
    assert stats["voip"].count == 2
    assert stats["hardware"].resolution_rate == 0.0


def test_averages():
    stats = load_tickets(SAMPLE)
    assert round(stats["network"].avg_hours_open, 2) == 18.62


def test_write_excel_rows(tmp_path):
    stats = load_tickets(SAMPLE)
    out = tmp_path / "r.xlsx"
    write_excel(stats, out)
    ws = load_workbook(out).active
    assert ws is not None
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0] == ("category", "tickets", "resolved", "resolution_rate", "avg_hours_open")
    assert len(rows) == 1 + len(stats)


def test_main_writes_both_files(tmp_path):
    assert main(["report", str(SAMPLE), str(tmp_path)]) == 0
    assert (tmp_path / "report.xlsx").stat().st_size > 0
    assert (tmp_path / "report.docx").stat().st_size > 0
