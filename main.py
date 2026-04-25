#!/usr/bin/env python3

import numbers
import ollama
import json
import os
from collections import defaultdict
import openpyxl
from openpyxl.utils import get_column_letter
import math
import csv
import io
from collections import Counter
from openpyxl.styles import Font
from pdf2image import convert_from_path


# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

model = os.getenv("MODEL", "frob/nuextract-2.0:latest")

PDF_PATH   = "input4.pdf"
PAGE_START = 4   # Fixed: was 5, page 4 has the first transactions
PAGE_END   = 6
OUTPUT_XLS = "transactions.xlsx"


# ─────────────────────────────────────────────
# ASK USER FOR COLUMN ORDER
# ─────────────────────────────────────────────

print("\n=== Column Order Helper ===")
print("Look at the Transaction Details table header in your PDF.")
print("Enter the column names in order, separated by commas.")
print("Example: Date, Category, Action, Symbol/CUSIP, Description, Quantity, Price/Rate per Share($), Charges/Interest($), Amount($), Realized Gain/Loss($)")
print()

raw_columns = input("Enter columns: ").strip()
column_names = [c.strip() for c in raw_columns.split(",")]

column_hint = (
    "The transaction table columns are, in order: "
    + ", ".join(f"{i+1}={name}" for i, name in enumerate(column_names))
    + ". "
    "Use the column named 'Amount($)' (or similar) for the dollar amount of each transaction, "
    "NOT the 'Price/Rate per Share' column. "
    "Use 'Realized Gain/Loss($)' (or similar) for realized_gain_loss. "
    "Quantity should come from the 'Quantity' column and will be negative for sales. "
    "IMPORTANT: some CSV cells contain multiple values separated by newline characters (\\n). "
    "Each newline-separated value belongs to a DIFFERENT row in the table. "
    "For example if the Amount($) cell reads '4.13\\n2.14\\n3.39\\n0.58', "
    "those are four separate amounts — one per transaction row — not one value. "
    "Match each value to its own row in order; do NOT copy the first value to every row."
)

print(f"\nColumn hint that will be sent to the model:\n  {column_hint}\n")


# ─────────────────────────────────────────────
# EXTRACTION TEMPLATE
# ─────────────────────────────────────────────

template = {
    "Statement Year": "string",
    "dividends": [
        {
            "date": "string",
            "cusip": "cusip or symbol",
            "description": "stock name",
            "amount": "number $",
            "action": "Reinvestment or Cash",
        }
    ],
    "purchases": [
        {
            "date": "string",
            "cusip": "cusip or symbol",
            "description": "stock name",
            "quantity": "number",
            "amount": "number $",
        }
    ],
    "sales": [
        {
            "date": "string",
            "cusip": "cusip or symbol",
            "description": "stock name",
            "quantity": "number",
            "amount": "number $",
            "realized_gain_loss": "number",
            "carry_value": "number or null",
        }
    ],
    "interest": [
        {
            "date": "string",
            "amount": "number $",
            "quantity": "number or null"
        }
    ]
}


# ─────────────────────────────────────────────
# PDF → TEXT HELPERS
# ─────────────────────────────────────────────

def img_to_rows(img_path, row_tolerance=8):
    """
    Run PaddleOCR on an image and reconstruct clean rows using bounding-box
    Y-coordinates instead of img2table's borderless table detector.

    img2table was merging adjacent rows into single cells because it tried to
    infer table structure from visual whitespace — and borderless financial
    statement tables fool it badly.

    This approach is simpler and far more reliable:
      1. Ask PaddleOCR for every text fragment + its bounding box.
      2. Group fragments whose vertical centres are within `row_tolerance`
         pixels of each other → these are on the same printed line.
      3. Within each group, sort by X so columns come out left-to-right.
      4. Join each row with a tab character and return one string per page.

    The result is plain, one-printed-line-per-text-line output with no merged
    cells, which the LLM can parse far more accurately.
    """
    from paddleocr import PaddleOCR as _PaddleOCR
    ocr = _PaddleOCR(lang="en")
    result = ocr.ocr(img_path, cls=True)

    if not result or not result[0]:
        return ""

    # Each entry: (bbox, (text, confidence))
    # bbox = [[x0,y0],[x1,y0],[x1,y1],[x0,y1]]
    fragments = []
    for line in result[0]:
        bbox, (text, conf) = line
        if conf < 0.5:
            continue
        y_center = (bbox[0][1] + bbox[2][1]) / 2
        x_left   = bbox[0][0]
        fragments.append((y_center, x_left, text))

    if not fragments:
        return ""

    # Sort top-to-bottom, then left-to-right within each row
    fragments.sort(key=lambda f: (f[0], f[1]))

    # Group into rows: a new row starts when y_center jumps by > row_tolerance
    rows = []
    current_row = [fragments[0]]
    for frag in fragments[1:]:
        if frag[0] - current_row[-1][0] > row_tolerance:
            rows.append(current_row)
            current_row = [frag]
        else:
            current_row.append(frag)
    rows.append(current_row)

    # Render each row as a tab-separated line
    lines = []
    for row in rows:
        row.sort(key=lambda f: f[1])           # left-to-right by x
        lines.append("\t".join(f[2] for f in row))

    return "\n".join(lines)


def pdf_page_to_text(pdf_path, page_number_start, page_number_end):
    images = convert_from_path(
        pdf_path,
        dpi=300,
        first_page=page_number_start,
        last_page=page_number_end,
        poppler_path="poppler/library/bin"
    )
    pages = []
    for i in range(page_number_end - page_number_start + 1):
        page_number = page_number_start + i
        temp_path = f"/tmp/page_{page_number}.png"
        images[i].save(temp_path, "PNG")
        pages.append(img_to_rows(temp_path))
    return pages


# ─────────────────────────────────────────────
# DATA HELPERS
# ─────────────────────────────────────────────

def merge_data(all_data, new_data):
    for key in all_data:
        all_data[key].extend(new_data.get(key, []))


def group_by_name(entries):
    grouped = defaultdict(list)
    for e in entries:
        grouped[e['description']].append(e)
    return grouped


def clean_data(data, page_data_list):
    """
    Post-processing fixes:

    1. Sales: reset carry_value if it looks like an exchange fee (< $1).
       NOTE: We no longer attempt to multiply amount × qty here — the column
       hint now makes the model return correct totals directly. Multiplying
       again would double-count and produce wildly wrong numbers.
    2. Dividends: remove any entries where amount == 0 (parsing artefacts).
    3. Interest: remove entries whose amount matches a known purchase OR
       dividend amount (misread rows from adjacent layout columns).
    4. Deduplicate: remove items that appear identically across pages.
    """

    # --- Fix 0: Correct swapped quantity ↔ price in sales/purchases ---
    # When the model confuses the Price/Rate column for Quantity, you get
    # a quantity like 551.2612 paired with an amount of 551.24 — implying
    # a per-share price of ~$1, which is impossible for any ETF here.
    # Heuristic: if amount / quantity < $10, the quantity is almost certainly
    # the per-share price; the real quantity is round(amount / quantity).
    for section in ('sales', 'purchases'):
        for item in data[section]:
            qty = abs(item.get('quantity') or 0)
            amt = abs(item.get('amount') or 0)
            if qty > 0 and amt > 0 and (amt / qty) < 10:
                corrected_qty = round(amt / qty)
                print(f"  [fix] {section} {item.get('cusip')}: quantity "
                      f"{qty} looks like a price → corrected to {corrected_qty} shares")
                item['quantity'] = corrected_qty

    # --- Fix 1: Sales carry_value — clear if it looks like an exchange fee ---
    for item in data['sales']:
        cv = item.get('carry_value') or 0
        if 0 < cv < 1.0:
            print(f"  [fix] Sale {item.get('cusip')}: clearing carry_value "
                  f"{cv} (looks like an exchange fee)")
            item['carry_value'] = None

    # --- Fix 2: Remove zero-amount dividends ---
    before = len(data['dividends'])
    data['dividends'] = [
        d for d in data['dividends']
        if float(d.get('amount') or 0) > 0
    ]
    removed = before - len(data['dividends'])
    if removed:
        print(f"  [fix] Removed {removed} zero-amount dividend(s)")

    # --- Fix 3: Remove interest entries that match purchase OR dividend amounts ---
    # Catches cases where an adjacent row (e.g. BSV dividend $7.80, VBR purchase
    # $189.52) bleeds into the interest section due to OCR layout confusion.
    purchase_amounts = {round(abs(float(p.get('amount', 0))), 2)
                        for p in data['purchases']}
    dividend_amounts = {round(abs(float(d.get('amount', 0))), 2)
                        for d in data['dividends']}
    known_non_interest = purchase_amounts | dividend_amounts
    before = len(data['interest'])
    data['interest'] = [
        i for i in data['interest']
        if round(abs(float(i.get('amount', 0))), 2) not in known_non_interest
    ]
    removed = before - len(data['interest'])
    if removed:
        print(f"  [fix] Removed {removed} interest entry/entries that matched "
              f"a purchase or dividend amount (likely misread rows)")

    # --- Fix 5: Deduplicate each section ---
    for section in data:
        seen = set()
        deduped = []
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


# ─────────────────────────────────────────────
# MAIN: EXTRACT
# ─────────────────────────────────────────────

if __name__ == "__main__":

    pages = pdf_page_to_text(PDF_PATH, PAGE_START, PAGE_END)

    all_page_data = []

    for i, page_text in enumerate(pages):
        page_num = PAGE_START + i
        print(f"\nSending page {page_num} to Ollama...")
        print(page_text)   # preview first 1000 chars so you can sanity-check it

        messages = [
            {"role": "template", "content": json.dumps(template)},
            {
                "role": "user",
                "content": (
                    f"IMPORTANT – column layout for this statement:\n{column_hint}\n\n"
                    "The text below is extracted line-by-line from a scanned PDF. "
                    "Each printed row is one text line; columns are separated by tabs. "
                    "When you extract sales, purchases, dividends and interest, "
                    "'amount' must be the TOTAL dollar value (the Amount($) column), "
                    "not the per-share price. There are NO multi-value cells — "
                    "every value on a line belongs only to that one transaction row."
                )
            },
            {
                "role": "assistant",
                "content": "Understood. I will use the Amount($) column for all 'amount' fields and treat each line as one transaction."
            },
            {
                "role": "user",
                "content": f"Page {page_num}:\n{page_text}"
            }
        ]

        response = ollama.chat(
            model=model,
            messages=messages,
            options={
                "temperature": 0,
                "num_predict": -2,
                "repeat_penalty": 1.1,
                "top_k": 1,
                "num_ctx": 8192*2
            }
        )

        try:
            page_data = json.loads(response.message.content)
            all_page_data.append(page_data)
        except Exception as e:
            print(f"Failed to parse page {page_num}: {e}")

    # ── Merge pages ──────────────────────────────────────────────────────────
    final_data = {
        "dividends": [],
        "purchases": [],
        "sales": [],
        "interest": []
    }

    for page_data in all_page_data:
        merge_data(final_data, page_data)

    # ── Post-processing fixes ────────────────────────────────────────────────
    print("\nRunning post-processing fixes...")
    data = clean_data(final_data, all_page_data)

# ── Normalize duplicate descriptions ─────────────────────────────
for section in ['purchases', 'sales', 'interest', 'dividends']:
    # Count descriptions
    desc_counts = Counter(
        (item.get('description') or '').strip()
        for item in data[section]
    )

    # Rename duplicates
    for item in data[section]:
        desc = (item.get('description') or '').strip()
        cusip = (item.get('cusip') or '').strip()

        if desc_counts[desc] > 1 and cusip:
            item['description'] = f"{desc}: {cusip}"

    # ── Sort ─────────────────────────────────────────────────────────────────
    for section in ['purchases', 'sales', 'interest']:
        data[section].sort(
            key=lambda x: (x.get('description') or '', x.get('date') or '')
        )
    for section in ['dividends']:
        data[section].sort(key=lambda x: (x.get('date') or ''))
    print(json.dumps(data, indent=2))


    # ─────────────────────────────────────────────
    # EXCEL OUTPUT
    # ─────────────────────────────────────────────

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    row = 1

    def write_row(cells, amount_cols=None, bold=False):
        global row
        if amount_cols is None:
            amount_cols = []
        for col, value in enumerate(cells, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            if col - 1 in amount_cols and (
                isinstance(value, (int, float)) or
                (isinstance(value, str) and value.startswith("="))
            ):
                cell.number_format = (
                    '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
                )
            if bold:
                cell.font = Font(bold=True)
        row += 1

    def add_space(rows):
        for _ in range(rows):
            write_row([])

    # ── Purchases ────────────────────────────────────────────────────────────
    write_row(['Purchases:'], bold=True)
    purchases_grouped = group_by_name(data['purchases'])
    for name, items in purchases_grouped.items():
        write_row([name], bold=True)
        add_space(1)
        write_row(['Date', 'Quantity', 'Amount'], bold=True)
        add_space(1)
        for item in items:
            write_row(
                [
                    item['date'],
                    f"{item['quantity']:.3f} shares" if item.get('quantity') else "",
                    math.fabs(float(item['amount']))
                ],
                amount_cols=[2]
            )
        add_space(1)
        write_row(
            ['', 'Total',
             f"=SUM({get_column_letter(3)}{row-len(items)-1}:{get_column_letter(3)}{row-1})"],
            amount_cols=[2, 3], bold=True
        )
        add_space(3)
    add_space(3)

    # ── Sales ────────────────────────────────────────────────────────────────
    write_row(['Sales:'], bold=True)
    sales_grouped = group_by_name(data['sales'])
    for name, items in sales_grouped.items():
        write_row([name], bold=True)
        add_space(1)
        write_row(
            ['Date', 'Quantity', 'Carry Value', 'Sales Price', 'Gain', 'Loss'],
            bold=True
        )
        add_space(1)
        for item in items:
            gain_loss   = item.get('realized_gain_loss', 0) or 0
            carry_value = item.get('carry_value')

            # Recalculate carry_value from amount − gain_loss if missing/tiny
            if (carry_value is None or carry_value < 0.1) and gain_loss is not None:
                carry_value = float(item['amount']) - float(gain_loss)

            write_row(
                [
                    item['date'],
                    f"{item['quantity']:.3f} shares" if item.get('quantity') else "",
                    float(carry_value) if carry_value is not None else None,
                    float(item['amount']) if item.get('amount') else None,
                    float(gain_loss) if gain_loss > 0 else 0,
                    float(-gain_loss) if gain_loss < 0 else 0
                ],
                amount_cols=[2, 3, 4, 5]
            )
        add_space(1)
        write_row(
            ['', 'Total',
             f"=SUM({get_column_letter(3)}{row-len(items)-1}:{get_column_letter(3)}{row-1})",
             f"=SUM({get_column_letter(4)}{row-len(items)-1}:{get_column_letter(4)}{row-1})",
             f"=SUM({get_column_letter(5)}{row-len(items)-1}:{get_column_letter(5)}{row-1})",
             f"=SUM({get_column_letter(6)}{row-len(items)-1}:{get_column_letter(6)}{row-1})"],
            amount_cols=[2, 3, 4, 5, 6], bold=True
        )
        add_space(3)

    # ── Dividends ────────────────────────────────────────────────────────────
    write_row(['Dividends:'], bold=True)
    write_row(['Date', 'Description', 'Amount'], bold=True)
    for d in data['dividends']:
        write_row([d['date'], d['description'], float(d['amount'])], amount_cols=[2])
    add_space(1)
    write_row(
        ['', 'Total',
         f"=SUM({get_column_letter(3)}{row-len(data['dividends'])-1}:{get_column_letter(3)}{row-1})"],
        amount_cols=[2, 3], bold=True
    )
    add_space(3)

    # ── Interest ─────────────────────────────────────────────────────────────
    write_row(['Interest:'], bold=True)
    write_row(['Date', 'Description', 'Amount'], bold=True)
    for i in data['interest']:
        write_row(
            [i['date'], "Interest", float(i['amount'])],
            amount_cols=[2], bold=True
        )
    add_space(1)
    write_row(
        ['', 'Total',
         f"=SUM({get_column_letter(3)}{row-len(data['interest'])-1}:{get_column_letter(3)}{row-1})"],
        amount_cols=[2, 3], bold=True
    )
    add_space(3)

    # ── Save ─────────────────────────────────────────────────────────────────
    wb.save(OUTPUT_XLS)
    print(f"\nSaved to {OUTPUT_XLS}")