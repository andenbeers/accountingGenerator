#!/usr/bin/env python3

import ollama
import json
import os
import base64
import math
from collections import defaultdict, Counter
from datetime import date
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from pdf2image import convert_from_path


# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

MODEL      = os.getenv("MODEL", "gemma3:27b")   # or gemma3:12b, qwen2.5vl:32b, etc.
PDF_PATH   = "input4.pdf"
PAGE_START = 4
PAGE_END   = 6
OUTPUT_XLS = "transactions.xlsx"
DPI        = 300


# ─────────────────────────────────────────────
# EXTRACTION TEMPLATE
# ─────────────────────────────────────────────

template = {
    "dividends": [
        {
            "date": "string (MM/DD format)",
            "cusip": "ticker symbol",
            "description": "fund name",
            "amount": "number",
            "action": "Cash or Reinvestment"
        }
    ],
    "purchases": [
        {
            "date": "string (MM/DD format)",
            "cusip": "ticker symbol",
            "description": "fund name",
            "quantity": "number (positive)",
            "amount": "number (total dollar amount from Amount$ column, NOT price per share)"
        }
    ],
    "sales": [
        {
            "date": "string (MM/DD format)",
            "cusip": "ticker symbol",
            "description": "fund name",
            "quantity": "number (positive)",
            "amount": "number (total dollar amount from Amount$ column)",
            "realized_gain_loss": "number (positive=gain, negative=loss)"
        }
    ],
    "interest": [
        {
            "date": "string (MM/DD format)",
            "description": "description text",
            "amount": "number"
        }
    ]
}

PROMPT = """You are extracting financial transactions from a brokerage statement image.

ONLY extract rows from the "Transaction Details" table section.
A valid transaction row has ALL of these: a Category (Dividend/Sale/Purchase/Interest), 
a Symbol/CUSIP, a Description, and an Amount($).

IGNORE everything that is NOT a transaction row:
- Page headers, account name, statement period
- Section titles like "Transaction Details", "Positions - Summary", etc.
- Column headers (Date, Category, Action, Symbol, Description, Quantity, Price, Amount)
- "Exchange Processing Fee $x.xx" lines
- "Total Transactions" line
- "Other Activity" lines
- Any table that is NOT the Transaction Details table (Positions, Cost Basis, Bank Sweep Activity, etc.)
- Footnotes, disclaimers, page numbers

CRITICAL RULES:
1. 'amount' must be the TOTAL dollar value from the Amount($) column — NOT Price/Rate per Share.
2. If a row has no date printed, inherit the most recent date printed above it in the table.
3. Each row is one transaction — do not merge rows.
4. For sales, realized_gain_loss is positive for a gain, negative for a loss.
5. Return ONLY valid JSON. No explanation, no markdown fences, no extra text.

Return exactly this structure:
""" + json.dumps(template, indent=2)


# ─────────────────────────────────────────────
# PDF → IMAGE HELPERS
# ─────────────────────────────────────────────

def pdf_to_images(pdf_path, page_start, page_end, dpi=DPI):
    """Convert PDF pages to PNG images, return list of (page_num, file_path)."""
    images = convert_from_path(
        pdf_path,
        dpi=dpi,
        first_page=page_start,
        last_page=page_end,
        poppler_path="poppler/library/bin"
    )
    result = []
    for i, image in enumerate(images):
        page_num  = page_start + i
        temp_path = f"/tmp/page_{page_num}.png"
        image.save(temp_path, "PNG")
        result.append((page_num, temp_path))
    return result


def image_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


# ─────────────────────────────────────────────
# MODEL CALL
# ─────────────────────────────────────────────

def call_vision_model(page_num, image_path):
    """Send one page image to the vision model, return parsed dict or None."""
    print(f"\nSending page {page_num} to {MODEL}...")

    img_b64 = image_to_base64(image_path)

    response = ollama.chat(
        model=MODEL,
        messages=[{
            "role": "user",
            "content": PROMPT,
            "images": [img_b64]
        }],
        options={
            "temperature": 0,
            "num_predict": 4096,
        }
    )

    raw = response.message.content.strip()

    # Strip markdown fences if the model added them anyway
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        data = json.loads(raw)
        total = sum(len(data.get(k, [])) for k in ['dividends', 'purchases', 'sales', 'interest'])
        print(f"  -> Extracted {total} entries from page {page_num}")
        for section in ['dividends', 'purchases', 'sales', 'interest']:
            items = data.get(section, [])
            if items:
                print(f"     {section}: {len(items)}")
        return data
    except Exception as e:
        print(f"  x Failed to parse page {page_num}: {e}")
        print(f"    Raw (first 500 chars): {raw[:500]}")
        return None


# ─────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────

def merge_data(base, new):
    for key in base:
        base[key].extend(new.get(key, []))


def group_by_name(entries):
    grouped = defaultdict(list)
    for e in entries:
        grouped[e['description']].append(e)
    return grouped


def clean_data(data):
    """
    Post-processing fixes:
    1. Correct swapped quantity/price in purchases and sales.
    2. Remove zero-amount dividends.
    3. Remove interest entries that match purchase or dividend amounts.
    4. Deduplicate across pages.
    """

    # Fix 1: Swapped quantity/price
    for section in ('sales', 'purchases'):
        for item in data[section]:
            qty = abs(float(item.get('quantity') or 0))
            amt = abs(float(item.get('amount') or 0))
            if qty > 0 and amt > 0 and (amt / qty) < 10:
                corrected = round(amt / qty)
                print(f"  [fix] {section} {item.get('cusip')}: qty {qty} "
                      f"looks like price/share, corrected to {corrected} shares")
                item['quantity'] = corrected

    # Fix 2: Zero-amount dividends
    before = len(data['dividends'])
    data['dividends'] = [d for d in data['dividends']
                         if float(d.get('amount') or 0) > 0]
    if len(data['dividends']) < before:
        print(f"  [fix] Removed {before - len(data['dividends'])} zero-amount dividend(s)")

    # Fix 3: Interest that matches purchase or dividend amounts
    purchase_amts = {round(abs(float(p.get('amount', 0))), 2) for p in data['purchases']}
    dividend_amts = {round(abs(float(d.get('amount', 0))), 2) for d in data['dividends']}
    non_interest  = purchase_amts | dividend_amts
    before = len(data['interest'])
    data['interest'] = [i for i in data['interest']
                        if round(abs(float(i.get('amount', 0))), 2) not in non_interest]
    if len(data['interest']) < before:
        print(f"  [fix] Removed {before - len(data['interest'])} misread interest entry/entries")

    # Fix 4: Deduplicate
    for section in data:
        seen, deduped = set(), []
        for item in data[section]:
            key = json.dumps(item, sort_keys=True)
            if key not in seen:
                seen.add(key)
                deduped.append(item)
        dupes = len(data[section]) - len(deduped)
        if dupes:
            print(f"  [fix] Removed {dupes} duplicate(s) from '{section}'")
        data[section] = deduped

    return data


def normalize_descriptions(data):
    """Append ticker to description when multiple entries share the same name."""
    for section in ['purchases', 'sales', 'interest', 'dividends']:
        counts = Counter((item.get('description') or '').strip()
                         for item in data[section])
        for item in data[section]:
            desc  = (item.get('description') or '').strip()
            cusip = (item.get('cusip') or '').strip()
            if counts[desc] > 1 and cusip:
                item['description'] = f"{desc}: {cusip}"


def sort_data(data):
    for section in ['purchases', 'sales', 'interest']:
        data[section].sort(
            key=lambda x: (x.get('description') or '', x.get('date') or '')
        )
    data['dividends'].sort(key=lambda x: x.get('date') or '')


# ─────────────────────────────────────────────
# VALIDATION
# ─────────────────────────────────────────────

# Update these from the Transactions Summary on page 4 of your statement
KNOWN_TOTALS = {
    "purchases": 2117.22,
    "sales":     2679.69,
    "dividends":   42.98,  # Cash Dividends only (excludes interest)
    "interest":     2.61,
}

def validate_totals(data):
    print("\n-- Validation ------------------------------------------")
    for section, expected in KNOWN_TOTALS.items():
        actual = sum(abs(float(item.get('amount', 0))) for item in data[section])
        diff   = abs(actual - expected)
        ok     = diff < 0.05
        status = "OK " if ok else "MISMATCH"
        print(f"  [{status}]  {section}: extracted ${actual:.2f}  "
              f"expected ${expected:.2f}"
              + (f"  (diff ${diff:.2f})" if not ok else ""))


# ─────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────

def write_excel(data, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    r = [1]  # mutable row counter

    AMOUNT_FMT = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'

    def wr(cells, amount_cols=None, bold=False):
        amount_cols = amount_cols or []
        for col, value in enumerate(cells, start=1):
            cell = ws.cell(row=r[0], column=col, value=value)
            if col - 1 in amount_cols and (
                isinstance(value, (int, float)) or
                (isinstance(value, str) and value.startswith("="))
            ):
                cell.number_format = AMOUNT_FMT
            if bold:
                cell.font = Font(bold=True)
        r[0] += 1

    def sp(n=1):
        r[0] += n

    def sumf(col_letter, r1, r2):
        return f"=SUM({col_letter}{r1}:{col_letter}{r2})"

    cl = get_column_letter

    # ── Purchases ──────────────────────────────────────────────────
    wr(['Purchases:'], bold=True)
    sp()
    for name, items in group_by_name(data['purchases']).items():
        wr([name], bold=True)
        sp()
        wr(['Date', 'Quantity', 'Amount'], bold=True)
        sp()
        r1 = r[0]
        for item in items:
            wr([
                item.get('date'),
                f"{float(item['quantity']):.3f} shares" if item.get('quantity') else "",
                math.fabs(float(item['amount']))
            ], amount_cols=[2])
        r2 = r[0] - 1
        sp()
        wr(['', 'Total', sumf('C', r1, r2)], amount_cols=[2], bold=True)
        sp(3)
    sp(2)

    # ── Sales ──────────────────────────────────────────────────────
    wr(['Sales:'], bold=True)
    sp()
    for name, items in group_by_name(data['sales']).items():
        wr([name], bold=True)
        sp()
        wr(['Date', 'Quantity', 'Carry Value', 'Sales Price', 'Gain', 'Loss'], bold=True)
        sp()
        r1 = r[0]
        for item in items:
            gl  = float(item.get('realized_gain_loss') or 0)
            sp_ = float(item.get('amount') or 0)
            cv  = sp_ - gl
            wr([
                item.get('date'),
                f"{float(item['quantity']):.3f} shares" if item.get('quantity') else "",
                cv,
                sp_,
                gl if gl > 0 else 0,
                -gl if gl < 0 else 0,
            ], amount_cols=[2, 3, 4, 5])
        r2 = r[0] - 1
        sp()
        wr(['', 'Total',
            sumf(cl(3), r1, r2), sumf(cl(4), r1, r2),
            sumf(cl(5), r1, r2), sumf(cl(6), r1, r2)],
           amount_cols=[2, 3, 4, 5], bold=True)
        sp(3)
    sp(2)

    # ── Dividends ──────────────────────────────────────────────────
    wr(['Dividends:'], bold=True)
    wr(['Date', 'Description', 'Amount'], bold=True)
    r1 = r[0]
    for d in data['dividends']:
        wr([d.get('date'), d.get('description'), float(d['amount'])], amount_cols=[2])
    r2 = r[0] - 1
    sp()
    wr(['', 'Total', sumf('C', r1, r2)], amount_cols=[2], bold=True)
    sp(3)

    # ── Interest ───────────────────────────────────────────────────
    wr(['Interest:'], bold=True)
    wr(['Date', 'Description', 'Amount'], bold=True)
    r1 = r[0]
    for i in data['interest']:
        wr([i.get('date'), i.get('description') or 'Interest', float(i['amount'])],
           amount_cols=[2])
    r2 = r[0] - 1
    sp()
    wr(['', 'Total', sumf('C', r1, r2)], amount_cols=[2], bold=True)

    # ── Column widths ──────────────────────────────────────────────
    for col, width in zip('ABCDEF', [14, 22, 16, 16, 14, 14]):
        ws.column_dimensions[col].width = width

    wb.save(output_path)
    print(f"\nSaved to {output_path}")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":

    # 1. Convert PDF pages to images
    print(f"Converting pages {PAGE_START}-{PAGE_END} of {PDF_PATH} to images...")
    page_images = pdf_to_images(PDF_PATH, PAGE_START, PAGE_END)

    # 2. Send each page image directly to the vision model
    all_page_data = []
    for page_num, image_path in page_images:
        page_data = call_vision_model(page_num, image_path)
        if page_data:
            all_page_data.append(page_data)

    if not all_page_data:
        print("No data extracted. Check model output above.")
        exit(1)

    # 3. Merge pages
    final_data = {"dividends": [], "purchases": [], "sales": [], "interest": []}
    for page_data in all_page_data:
        merge_data(final_data, page_data)

    # 4. Clean, normalize, sort
    print("\nRunning post-processing...")
    data = clean_data(final_data)
    normalize_descriptions(data)
    sort_data(data)

    # 5. Validate against statement totals
    validate_totals(data)

    # 6. Print extracted JSON
    print("\n-- Extracted Data --------------------------------------")
    print(json.dumps(data, indent=2))

    # 7. Write Excel
    write_excel(data, OUTPUT_XLS)