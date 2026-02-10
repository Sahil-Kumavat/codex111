# Personal Expense & Loan Tracker

This project provides a simple command-line program to record daily expenses and friend loans in an Excel workbook, plus a Power BI-friendly sheet for analysis.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Initialize the workbook

```bash
python expense_tracker.py init
```

This creates `data/finance_tracker.xlsx` with three sheets:
- `Expenses`
- `Loans`
- `PowerBI` (tabular sheet ready for Power BI)

## Add an expense

```bash
python expense_tracker.py add-expense \
  --date 2024-05-12 \
  --category Groceries \
  --amount 25.5 \
  --payment-method Card \
  --merchant "Local Market" \
  --notes "Fruits and vegetables"
```

## Add a loan

```bash
python expense_tracker.py add-loan \
  --date 2024-05-13 \
  --person "Amit" \
  --direction lent \
  --amount 100 \
  --status open \
  --due-date 2024-06-01 \
  --notes "Dinner split"
```

## Update the Power BI sheet

```bash
python expense_tracker.py powerbi
```

## Get a quick summary

```bash
python expense_tracker.py summary
```

## Notes
- The `PowerBI` sheet stores a single tabular view of all expenses and loans, which makes it easy to load into Power BI.
- Dates accept `YYYY-MM-DD`, `DD-MM-YYYY`, `DD/MM/YYYY`, `MM/DD/YYYY`, or `today`.
