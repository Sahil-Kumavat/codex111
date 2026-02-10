#!/usr/bin/env python3
"""Personal finance tracker for daily expenses and friend loans."""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Literal

import pandas as pd

DEFAULT_WORKBOOK = Path("data") / "finance_tracker.xlsx"

EXPENSE_COLUMNS = [
    "Date",
    "Category",
    "Amount",
    "PaymentMethod",
    "Merchant",
    "Notes",
]

LOAN_COLUMNS = [
    "Date",
    "Person",
    "Direction",
    "Amount",
    "Status",
    "DueDate",
    "Notes",
]

POWERBI_COLUMNS = [
    "RecordType",
    "Date",
    "Counterparty",
    "Category",
    "Amount",
    "Direction",
    "Status",
    "Notes",
]


def parse_date(value: str) -> str:
    """Parse a date input and normalize to YYYY-MM-DD string."""
    if value.lower() == "today":
        return date.today().isoformat()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(
        "Date must be YYYY-MM-DD, DD-MM-YYYY, DD/MM/YYYY, MM/DD/YYYY, or 'today'."
    )


@dataclass
class WorkbookManager:
    path: Path

    def ensure_workbook(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.exists():
            return
        with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
            pd.DataFrame(columns=EXPENSE_COLUMNS).to_excel(
                writer, sheet_name="Expenses", index=False
            )
            pd.DataFrame(columns=LOAN_COLUMNS).to_excel(
                writer, sheet_name="Loans", index=False
            )
            pd.DataFrame(columns=POWERBI_COLUMNS).to_excel(
                writer, sheet_name="PowerBI", index=False
            )

    def read_sheet(self, sheet_name: str, columns: list[str]) -> pd.DataFrame:
        self.ensure_workbook()
        try:
            return pd.read_excel(self.path, sheet_name=sheet_name)
        except ValueError:
            return pd.DataFrame(columns=columns)

    def write_sheet(self, sheet_name: str, dataframe: pd.DataFrame) -> None:
        self.ensure_workbook()
        with pd.ExcelWriter(
            self.path, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)


@dataclass
class ExpenseEntry:
    date: str
    category: str
    amount: float
    payment_method: str
    merchant: str
    notes: str

    def to_row(self) -> dict[str, object]:
        return {
            "Date": self.date,
            "Category": self.category,
            "Amount": self.amount,
            "PaymentMethod": self.payment_method,
            "Merchant": self.merchant,
            "Notes": self.notes,
        }


@dataclass
class LoanEntry:
    date: str
    person: str
    direction: Literal["borrowed", "lent"]
    amount: float
    status: str
    due_date: str
    notes: str

    def to_row(self) -> dict[str, object]:
        return {
            "Date": self.date,
            "Person": self.person,
            "Direction": self.direction,
            "Amount": self.amount,
            "Status": self.status,
            "DueDate": self.due_date,
            "Notes": self.notes,
        }


def add_expense(manager: WorkbookManager, entry: ExpenseEntry) -> None:
    expenses = manager.read_sheet("Expenses", EXPENSE_COLUMNS)
    expenses = pd.concat([expenses, pd.DataFrame([entry.to_row()])], ignore_index=True)
    manager.write_sheet("Expenses", expenses)


def add_loan(manager: WorkbookManager, entry: LoanEntry) -> None:
    loans = manager.read_sheet("Loans", LOAN_COLUMNS)
    loans = pd.concat([loans, pd.DataFrame([entry.to_row()])], ignore_index=True)
    manager.write_sheet("Loans", loans)


def update_powerbi(manager: WorkbookManager) -> None:
    expenses = manager.read_sheet("Expenses", EXPENSE_COLUMNS)
    loans = manager.read_sheet("Loans", LOAN_COLUMNS)

    expense_rows = pd.DataFrame(
        {
            "RecordType": "Expense",
            "Date": expenses.get("Date"),
            "Counterparty": expenses.get("Merchant"),
            "Category": expenses.get("Category"),
            "Amount": expenses.get("Amount"),
            "Direction": "outflow",
            "Status": "n/a",
            "Notes": expenses.get("Notes"),
        }
    )

    loan_rows = pd.DataFrame(
        {
            "RecordType": "Loan",
            "Date": loans.get("Date"),
            "Counterparty": loans.get("Person"),
            "Category": "loan",
            "Amount": loans.get("Amount"),
            "Direction": loans.get("Direction"),
            "Status": loans.get("Status"),
            "Notes": loans.get("Notes"),
        }
    )

    combined = pd.concat([expense_rows, loan_rows], ignore_index=True).fillna("")
    manager.write_sheet("PowerBI", combined)


def summary(manager: WorkbookManager) -> None:
    expenses = manager.read_sheet("Expenses", EXPENSE_COLUMNS)
    loans = manager.read_sheet("Loans", LOAN_COLUMNS)

    if not expenses.empty:
        expenses["Month"] = pd.to_datetime(expenses["Date"]).dt.to_period("M")
        by_category = expenses.groupby("Category")["Amount"].sum().sort_values(ascending=False)
        by_month = expenses.groupby("Month")["Amount"].sum().sort_index()
        print("\nExpense totals by category:")
        print(by_category.to_string())
        print("\nExpense totals by month:")
        print(by_month.to_string())
    else:
        print("\nNo expenses recorded yet.")

    if not loans.empty:
        net = loans.assign(
            SignedAmount=loans.apply(
                lambda row: row["Amount"] if row["Direction"] == "lent" else -row["Amount"],
                axis=1,
            )
        )
        balance = net["SignedAmount"].sum()
        print("\nLoan balance (positive = others owe you):")
        print(balance)
    else:
        print("\nNo loans recorded yet.")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Track daily expenses and friend loans in an Excel workbook."
    )
    parser.add_argument(
        "--workbook",
        default=str(DEFAULT_WORKBOOK),
        help="Path to the Excel workbook (default: data/finance_tracker.xlsx)",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("init", help="Create the workbook if it does not exist.")

    expense_parser = subparsers.add_parser("add-expense", help="Record a new expense.")
    expense_parser.add_argument("--date", type=parse_date, required=True)
    expense_parser.add_argument("--category", required=True)
    expense_parser.add_argument("--amount", type=float, required=True)
    expense_parser.add_argument("--payment-method", default="cash")
    expense_parser.add_argument("--merchant", default="")
    expense_parser.add_argument("--notes", default="")

    loan_parser = subparsers.add_parser("add-loan", help="Record a loan with a friend.")
    loan_parser.add_argument("--date", type=parse_date, required=True)
    loan_parser.add_argument("--person", required=True)
    loan_parser.add_argument("--direction", choices=["borrowed", "lent"], required=True)
    loan_parser.add_argument("--amount", type=float, required=True)
    loan_parser.add_argument("--status", default="open")
    loan_parser.add_argument("--due-date", type=parse_date, default="")
    loan_parser.add_argument("--notes", default="")

    subparsers.add_parser("summary", help="Print summary totals for expenses and loans.")
    subparsers.add_parser(
        "powerbi", help="Update the PowerBI sheet for easier BI ingestion."
    )

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    manager = WorkbookManager(Path(args.workbook))
    manager.ensure_workbook()

    if args.command == "init":
        print(f"Workbook ready at {manager.path}")
        return

    if args.command == "add-expense":
        entry = ExpenseEntry(
            date=args.date,
            category=args.category,
            amount=args.amount,
            payment_method=args.payment_method,
            merchant=args.merchant,
            notes=args.notes,
        )
        add_expense(manager, entry)
        update_powerbi(manager)
        print("Expense recorded.")
        return

    if args.command == "add-loan":
        entry = LoanEntry(
            date=args.date,
            person=args.person,
            direction=args.direction,
            amount=args.amount,
            status=args.status,
            due_date=args.due_date,
            notes=args.notes,
        )
        add_loan(manager, entry)
        update_powerbi(manager)
        print("Loan recorded.")
        return

    if args.command == "summary":
        summary(manager)
        return

    if args.command == "powerbi":
        update_powerbi(manager)
        print("PowerBI sheet updated.")
        return


if __name__ == "__main__":
    main()
