#!/usr/bin/env python3
"""
PDF → docstrange (markdown) → nuextract (Ollama) → Excel
"""

import json
import os
import math
from collections import defaultdict

os.environ["PATH"] = os.path.abspath(r"poppler\library\bin") + ";" + os.environ["PATH"]

from docstrange import DocumentExtractor
from pypdf import PdfReader, PdfWriter
import ollama
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ── CONFIG ────────────────────────────────────────────────────────────────────

MODEL      = os.getenv("MODEL", "frob/nuextract-2.0:latest")
PDF_PATH   = "input4.pdf"
PAGE_START = 4   # 1-indexed, inclusive
PAGE_END   = 6   # 1-indexed, inclusive
OUTPUT_XLS = "transactions.xlsx"

extractor = DocumentExtractor(gpu=True)

# ── EXTRACTION TEMPLATE ───────────────────────────────────────────────────────

TEMPLATE = {
    "dividends": [
        {"date": "string", "cusip": "string", "description": "string",
         "amount": "number", "action": "Reinvestment or Cash"}
    ],
    "purchases": [
        {"date": "string", "cusip": "string", "description": "string",
         "quantity": "number", "amount": "number"}
    ],
    "sales": [
        {"date": "string", "cusip": "string", "description": "string",
         "quantity": "number", "amount": "number",
         "realized_gain_loss": "number", "carry_value": "number"}
    ],
    # description required — anchors nuextract to real interest rows
    # (e.g. "BANK INT 081624-091524") and prevents it grabbing
    # numeric amounts from the Bank Sweep Activity section
    "interest": [
        {"date": "string", "description": "string", "amount": "number"}
    ],
}

USER_PROMPT = (
    "Extract transactions ONLY from the 'Transaction Details' table. "
    "IGNORE all other sections: 'Bank Sweep Activity', 'Cost Basis Lot Details', "
    "'Positions', 'Account Summary', and any totals rows. "
    "Use the Amount($) column for all 'amount' fields. "
    "Each table row is one transaction. Quantity is negative for sales. "
    "Only add a row to 'interest' when the Category column explicitly says 'Interest' — "
    "bank sweep credits, beginning/ending balances, and brokerage credits are NOT interest."
)

ASSISTANT_ACK = (
    "Understood. I will only extract rows from the Transaction Details table, "
    "use Amount($) for all amounts, and only add rows to 'interest' "
    "when Category is explicitly 'Interest'."
)


# ── STEP 1: PDF → MARKDOWN ────────────────────────────────────────────────────

def pdf_page_to_markdown(pdf_path: str, page_idx: int) -> str:
    writer = PdfWriter()
    writer.add_page(PdfReader(pdf_path).pages[page_idx])
    tmp = f"/tmp/_page_{page_idx + 1}.pdf"
    with open(tmp, "wb") as f:
        writer.write(f)
    return extractor.extract(tmp).extract_markdown()


def extract_pages(pdf_path: str, start: int, end: int) -> list[tuple[int, str]]:
    pages = []
    for idx in range(start - 1, end):
        print(f"  docstrange → page {idx + 1} …", end=" ", flush=True)
        md = pdf_page_to_markdown(pdf_path, idx)
        print(f"done ({len(md)} chars)")
        print(f"  preview: {md[:300].strip()!r}\n")
        if md.strip():
            pages.append((idx + 1, md))
        else:
            print(f"  [warn] page {idx + 1} empty, skipping")
    return pages


# ── STEP 2: MARKDOWN → JSON (nuextract via Ollama) ────────────────────────────

def extract_transactions(page_num: int, markdown: str) -> dict | None:
    messages = [
        {"role": "template",  "content": json.dumps(TEMPLATE)},
        {"role": "user",      "content": USER_PROMPT},
        {"role": "assistant", "content": ASSISTANT_ACK},
        {"role": "user",      "content": f"Page {page_num}:\n{markdown}"},
    ]
    response = ollama.chat(
        model=MODEL,
        messages=messages,
        options={"temperature": 0, "num_predict": -2, "top_k": 1, "num_ctx": 16384},
        
    )
    try:
        return json.loads(response.message.content)
    except Exception as e:
        print(f"  [warn] page {page_num} parse failed: {e}")
        return None


def merge(base: dict, new: dict) -> None:
    for key in base:
        base[key].extend(new.get(key, []))


def deduplicate(data: dict) -> dict:
    for section, items in data.items():
        seen, out = set(), []
        for item in items:
            k = json.dumps(item, sort_keys=True)
            if k not in seen:
                seen.add(k)
                out.append(item)
        removed = len(items) - len(out)
        if removed:
            print(f"  [fix] removed {removed} duplicate(s) from '{section}'")
        data[section] = out

    # Interest-specific filters:
    # 1. Drop entries with no day in the date (e.g. "2024-09") — those are summary totals
    # 2. Drop entries where description contains "rate" — those are interest rate percentages, not payments
    # 3. Deduplicate by amount — Bank Sweep Activity echoes the same payment with a different date
    clean = []
    for item in data["interest"]:
        date = str(item.get("date") or "")
        amt  = round(float(item.get("amount") or 0), 2)
        if date.count("-") < 2:
            print(f"  [fix] dropped interest row — date has no day: {date}")
            continue
        if "rate" in str(item.get("description") or "").lower():
            print(f"  [fix] dropped interest row — description looks like a rate: {item.get('description')}")
            continue
        clean.append(item)

    seen_amt, deduped = set(), []
    for item in clean:
        amt = round(float(item.get("amount") or 0), 2)
        if amt not in seen_amt:
            seen_amt.add(amt)
            deduped.append(item)
    removed = len(data["interest"]) - len(deduped)
    if removed:
        print(f"  [fix] interest: kept {len(deduped)} of {len(data['interest'])} entries")
    data["interest"] = deduped
    return data


# ── STEP 3: EXCEL OUTPUT ──────────────────────────────────────────────────────

MONEY = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

class Sheet:
    def __init__(self):
        self.wb  = openpyxl.Workbook()
        self.ws  = self.wb.active
        self.ws.title = "Transactions"
        self.row = 1

    def write(self, cells, money_cols=(), bold=False):
        money_cols = set(money_cols)
        for col, val in enumerate(cells, 1):
            c = self.ws.cell(row=self.row, column=col, value=val)
            is_money = isinstance(val, (int, float)) or (isinstance(val, str) and val.startswith('='))
            if (col - 1) in money_cols and is_money:
                c.number_format = MONEY
            if bold:
                c.font = Font(bold=True)
        self.row += 1

    def blank(self, n=1):
        self.row += n

    def sum_col(self, col: int, r1: int, r2: int) -> str:
        return f"=SUM({get_column_letter(col)}{r1}:{get_column_letter(col)}{r2})"

    def save(self, path: str):
        self.wb.save(path)
        print(f"Saved → {path}")

    def purchases(self, items: list):
        if not items:
            return
        self.write(["Purchases"], bold=True)
        for name, rows in _group(items, "description").items():
            self.blank()
            self.write([name], bold=True)
            self.write(["Date", "Quantity", "Amount"], bold=True)
            r0 = self.row
            for r in rows:
                self.write([r.get("date"), _shares(r), _abs(r, "amount")], money_cols=[2])
            r1 = self.row - 1
            self.write(["", "Total", self.sum_col(3, r0, r1)], money_cols=[2], bold=True)
        self.blank(2)

    def sales(self, items: list):
        if not items:
            return
        self.write(["Sales"], bold=True)
        for name, rows in _group(items, "description").items():
            self.blank()
            self.write([name], bold=True)
            self.write(["Date", "Quantity", "Carry Value", "Sale Amount", "Gain", "Loss"], bold=True)
            r0 = self.row
            for r in rows:
                gl  = float(r.get("realized_gain_loss") or 0)
                amt = float(r.get("amount") or 0)
                cv  = float(r.get("carry_value") or 0) or (amt - gl)
                self.write(
                    [r.get("date"), _shares(r), cv, amt,
                     gl if gl > 0 else 0, -gl if gl < 0 else 0],
                    money_cols=[2, 3, 4, 5],
                )
            r1 = self.row - 1
            self.write(
                ["", "Total",
                 self.sum_col(3, r0, r1), self.sum_col(4, r0, r1),
                 self.sum_col(5, r0, r1), self.sum_col(6, r0, r1)],
                money_cols=[2, 3, 4, 5], bold=True,
            )
        self.blank(2)

    def dividends(self, items: list):
        if not items:
            return
        self.write(["Dividends"], bold=True)
        self.write(["Date", "Description", "Amount"], bold=True)
        r0 = self.row
        for r in items:
            self.write([r.get("date"), r.get("description"), _abs(r, "amount")], money_cols=[2])
        r1 = self.row - 1
        self.write(["", "Total", self.sum_col(3, r0, r1)], money_cols=[2], bold=True)
        self.blank(2)

    def interest(self, items: list):
        if not items:
            return
        self.write(["Interest"], bold=True)
        self.write(["Date", "Description", "Amount"], bold=True)
        r0 = self.row
        for r in items:
            self.write([r.get("date"), r.get("description", ""), _abs(r, "amount")], money_cols=[2])
        r1 = self.row - 1
        self.write(["", "Total", self.sum_col(3, r0, r1)], money_cols=[2], bold=True)


# ── HELPERS ───────────────────────────────────────────────────────────────────

_ticker_cache: dict[str, str] = {}

def _resolve_name(cusip: str | None, fallback: str | None) -> str:
    """Use yfinance name as-is if found; otherwise title-case the statement description."""
    symbol = (cusip or "").strip().upper()
    if symbol and symbol not in _ticker_cache:
        try:
            import yfinance as yf
            info = yf.Ticker(symbol).info
            name = (info.get("shortName") or info.get("longName") or "").strip()
            _ticker_cache[symbol] = name or _title(fallback or symbol)
            print(f"  [ticker] {symbol} → {_ticker_cache[symbol]}")
        except Exception:
            _ticker_cache[symbol] = _title(fallback or symbol)
    return _ticker_cache.get(symbol) or _title(fallback or "Unknown")

def _title(s: str) -> str:
    """Title-case: first letter of each word capitalised, rest lower."""
    return s.strip().title() if s else s

def resolve_names(data: dict) -> dict:
    """Enrich every item's description using its cusip/ticker symbol."""
    for section in data.values():
        for item in section:
            item["description"] = _resolve_name(item.get("cusip"), item.get("description"))
    return data

def _group(items: list, key: str) -> dict:
    out = defaultdict(list)
    for item in items:
        out[item.get(key) or "Unknown"].append(item)
    return out

def _abs(item: dict, key: str) -> float:
    return math.fabs(float(item.get(key) or 0))

def _shares(item: dict) -> str:
    q = item.get("quantity")
    return f"{abs(float(q)):.3f} shares" if q else ""


# ── MAIN ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        ollama.generate(model=MODEL, prompt="", keep_alive=0)
        print("[0] Ollama model unloaded from VRAM")
    except Exception:
        pass
    print(f"\n[1] Extracting pages {PAGE_START}–{PAGE_END} with docstrange …")
    pages = extract_pages(PDF_PATH, PAGE_START, PAGE_END)

    print(f"\n[2] Sending to Ollama ({MODEL}) …")
    data: dict = {"dividends": [], "purchases": [], "sales": [], "interest": []}
    for page_num, markdown in pages:
        print(f"  page {page_num} …")
        result = extract_transactions(page_num, markdown)
        if result:
            merge(data, result)

    print("\n[3] Deduplicating …")
    data = deduplicate(data)
    print(json.dumps(data, indent=2))

    print("\n[4] Resolving ticker names …")
    data = resolve_names(data)

    print("\n[5] Writing Excel …")
    sheet = Sheet()
    sheet.purchases(data["purchases"])
    sheet.sales(data["sales"])
    sheet.dividends(data["dividends"])
    sheet.interest(data["interest"])
    sheet.save(OUTPUT_XLS)