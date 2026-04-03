 #!/usr/bin/env python3

import numbers

import ollama
import json
import os
from collections import defaultdict
from paddleocr import PaddleOCR
import pandas as pd
import pdfplumber
import ocrmypdf
import openpyxl
from openpyxl.utils import get_column_letter
import math
import numpy as np


model = os.getenv("MODEL", "frob/nuextract-2.0:latest")

def group_by_name(entries):
    grouped = defaultdict(list)
    for e in entries:
        grouped[e['description']].append(e)
    return grouped


template = {
  "Statement Year": "string",
  "dividends": [
    {
      "date": "string",
      "description": "stock name",
      "amount": "number $",
      "action": "Reinvestment or Cash"
    }
  ],
  "purchases": [
    {
      "date": "string",
      "description": "stock name",
      "quantity": "number",
      "amount": "number $",
    }
  ],
  "sales": [
    {
      "date": "string",
      "description": "stock name",
      "quantity": "number",
      "amount": "number $",
      "realized_gain_loss": "number",
      "carry_value": "number",
    }
  ],
  "interest": [
      {
        "date": "string",
        "amount": "number $"
      }
  ]
}

examples = [
  {
    "input": "Date, Category, Action, Symbol, Description, Quanity, Price/Rate, Charged Interest, Amount ($), Realized Gain (Loss)($)"
    "9/6, Dividend, Cash Dividend, IBTF, ISHARES IBONDS TERM, , , , 17.33,",
    "output": """{
      "dividends": [
        {
          "date": "09/06",
          "description": "ISHARES IBONDS TERM",
          "amount": 17.33,
          "action": "Cash Dividend"
        }
      ],
      "purchases": [],
      "sales": []
    }"""
  },
  {
    "input": "Date, Category, Action, Symbol, Description, Quanity, Price/Rate, Charged Interest, Amount ($), Realized Gain (Loss)($)"
    "9/11, Purchase, Buy, VUG, VANGUARD GROWTH ETF, 6.000, 363.0799, , -2178.48,",
    "output": """{
      "dividends": [],
      "purchases": [
        {
          "date": "09/11",
          "description": "VANGUARD GROWTH ETF",
          "quantity": 6.0,
          "amount": -2178.48
        }
      ],
      "sales": []
    }"""
  },
  {
    "input": "Date, Category, Action, Symbol, Description, Quanity, Price/Rate, Charged Interest, Amount ($), Realized Gain (Loss)($)"
    "9/11, Sale, Sell, IBTE, ISHARES IBONDS TERM TREASURY ETF, 331.000, 23.9300, 0.27, 7920.56, 12.97",
    "output": """{
      "dividends": [],
      "purchases": [],
      "sales": [
        {
          "date": "09/11",
          "description": "ISHARES IBONDS TERM TREASURY ETF",
          "quantity": 331.0,
          "amount": 7920.56,
          "realized_gain_loss": 12.97
        }
      ]
    }"""
  }
]
  
from pdf2image import convert_from_path

POPPLER_PATH = r"poppler\Library\bin"  

def ocr_pdf_to_pages(pdf_path):
    images = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path=POPPLER_PATH,
    )

    ocr = PaddleOCR(
        use_angle_cls=True,
        lang='en',
        det_db_box_thresh=0.6,
        det_db_unclip_ratio=2.0,
        det_limit_side_len=2000,
        ocr_version='PP-OCRv4',
    )

    pages_text = []

    for i, img in enumerate(images):
        print(f"Processing page {i+1}...")

        img_path = f"temp_page_{i}.png"
        img.save(img_path)

        result = ocr.ocr(img_path)

        all_words = []

        for page in result:
            if 'rec_texts' in page and 'dt_polys' in page:
                for text, poly in zip(page['rec_texts'], page['dt_polys']):
                    y_bottom = np.max(poly[:, 1])
                    x_center = np.mean(poly[:, 0])
                    all_words.append((y_bottom, x_center, text))

        all_words.sort(key=lambda x: x[0])

        lines = []
        current_line = []
        last_y = None

        for y, x, text in all_words:
            y_thresh = 10
            if last_y is None or abs(y - last_y) <= y_thresh:
                current_line.append((x, text))
            else:
                current_line.sort(key=lambda w: w[0])
                lines.append(", ".join(w[1] for w in current_line))
                current_line = [(x, text)]
            last_y = y

        if current_line:
            current_line.sort(key=lambda w: w[0])
            lines.append(", ".join(w[1] for w in current_line))

        page_text = "\n".join(lines)
        pages_text.append(page_text)

    return pages_text

def merge_data(all_data, new_data):
    for key in all_data:
        all_data[key].extend(new_data.get(key, []))


if __name__ == "__main__":

    pages = ocr_pdf_to_pages("input.pdf")
    print(pages)

    # Store each page's parsed JSON
    all_page_data = []

    for i, page in enumerate(pages):
        print(f"Sending page {i+1} to Ollama...")

        messages = [
            {"role": "template", "content": json.dumps(template)}
        ]

        for ex in examples:
            messages.append({"role": "examples.input", "content": ex["input"]})
            messages.append({"role": "examples.output", "content": ex["output"]})

        messages.append({
            "role": "user",
            "content": f"Page {i+1}:\n{page}"
        })

        response = ollama.chat(
            model=model,
            messages=messages,
            options={"temperature": 0.2}
        )

        try:
            page_data = json.loads(response.message.content)

            

            all_page_data.append(page_data)  # ✅ store it
        except Exception as e:
            print(f"Failed to parse page {i+1}: {e}")

    # 🔗 Merge AFTER loop
    final_data = {
        "dividends": [],
        "purchases": [],
        "sales": [],
        "interest": []
    }

    for page_data in all_page_data:
        merge_data(final_data, page_data)

    data = final_data

    # we need to sort the final data, each section of it by name of ['description'] and then by date, to make it easier to read in the excel file
    for section in ['dividends', 'purchases', 'sales', 'interest']:
      data[section].sort(key=lambda x: (x.get('description',''), x.get('date','')))

    print(data)


# Create a new workbook and select the active sheet


    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    row = 1

    def write_row(cells, amount_cols=None):
        """Write a row to the sheet.
        amount_cols: list of zero-based column indices to format as accounting
        """
        global row
        if amount_cols is None:
            amount_cols = []
        for col, value in enumerate(cells, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            if col-1 in amount_cols and isinstance(value, (int, float)):
                cell.number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'  # <-- proper accounting format string
        row += 1

    # Purchases
    write_row(['Purchases:'])
    purchases_grouped = group_by_name(data['purchases'])
    for name, items in purchases_grouped.items():
        write_row([name])
        write_row(['Date', 'Quantity', 'Amount'])
        for item in items:
            # Amount is column 2 (0-based index)
            write_row([item['date'], f"{item['quantity']:.3f} shares", math.fabs(float(item['amount']))], amount_cols=[2])
        row += 1

    # Sales
    write_row(['Sales:'])
    sales_grouped = group_by_name(data['sales'])
    for name, items in sales_grouped.items():
        write_row([name])
        write_row(['Date', 'Quantity', 'Carry Value', 'Sales Price', 'Gain', 'Loss'])
        for item in items:
            gain_loss = item.get('realized_gain_loss', 0)
            carry_value = item.get('carry_value')
            if carry_value is None and gain_loss is not None:
                carry_value = item['amount'] - gain_loss
            # Columns 2=Amount, 3=Carry_Value, 4=Gain, 5=Loss
            write_row([
                item['date'],
                f"{item['quantity']:.3f} shares",
                float(carry_value) if carry_value is not None else None,
                float(item['amount']),
                float(gain_loss) if gain_loss > 0 else None,
                float(-gain_loss) if gain_loss < 0 else None
            ], amount_cols=[2,3,4,5])
        row += 1

    # Dividends
    write_row(['Dividends:'])
    write_row(['Date', 'Description', 'Amount'])
    for d in data['dividends']:
        write_row([d['date'], d['description'], float(d['amount'])], amount_cols=[2])

    # Interest
    write_row(['Interest:'])
    write_row(['Date','Amount'])
    for i in data['interest']:
        write_row([i['date'], float(i['amount'])], amount_cols=[1])

    # Save workbook
    wb.save('transactions.xlsx')