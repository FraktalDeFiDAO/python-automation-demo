# python-automation-demo

A small, complete example of the kind of Python automation Copper Tone Technologies ships: read a CSV, aggregate it, and produce an Excel workbook plus a Word report. No network, no secrets, no paid services. Input in, files out.

## Run it

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl python-docx pytest
.venv/bin/python -m src/report                    # uses data/sample-tickets.csv, writes out/
.venv/bin/python -m src/report your.csv outdir    # your own CSV
```

The CSV needs columns `ticket_id, opened, category, priority, hours_open, resolved`. `src/report.py` loads it, aggregates per category (count, resolution rate, average hours open), writes `report.xlsx` (openpyxl) and `report.docx` (python-docx), and names the slowest category in plain language.

## Tests

```bash
.venv/bin/python -m pytest -q
```

Covers the aggregation, the averages, the Excel layout and the end-to-end run.

## What this pattern looks like in real work

Same shape, different sources: an API or a database instead of a CSV, email delivery of the report, a schedule instead of a manual run. Those parts are deliberately left out here so the example runs anywhere.
