 #!/usr/bin/env python3

import numbers

import ollama
import json
import os
from collections import defaultdict
from paddleocr import PaddleOCR
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import math
import numpy as np
from img2table.document import PDF
from img2table.ocr import PaddleOCR
import csv
import io
from openpyxl.styles import Font



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

  

from img2table.document import Image


def img_to_csv(img_path):
    ocr = PaddleOCR(lang="en")
    doc = Image(src=img_path)

    tables = doc.extract_tables(
        ocr=ocr,
        implicit_rows=True,
        borderless_tables=True,
        min_confidence=70,
    )

    table_csvs = []

    for table in tables:
        output = io.StringIO()
        writer = csv.writer(output)
        for row in table.content.values():
            writer.writerow([cell.value if cell else "" for cell in row])
        table_csvs.append(output.getvalue())

    return table_csvs


from pdf2image import convert_from_path

def pdf_page_to_csv(pdf_path, page_number_start, page_number_end):

    images = convert_from_path(
        pdf_path,
        dpi=300,
        first_page=page_number_start,
        last_page=page_number_end,
        poppler_path="poppler/library/bin"
    )
    pages = []
    # Save the rasterized page as a temp image
    for i in range(page_number_end - page_number_start + 1):
        page_number = page_number_start + i
        temp_path = f"/tmp/page_{page_number}.png"
        images[i].save(temp_path, "PNG")
        
        pages.append(img_to_csv(temp_path))
        
        # Reuse your working image pipeline
    return pages


def merge_data(all_data, new_data):
    for key in all_data:
        all_data[key].extend(new_data.get(key, []))




if __name__ == "__main__":
    pages = pdf_page_to_csv("input4.pdf", 5, 6)  # Example: get page 1 as CSV. Loop this for all pages you want to process.

    # Store each page's parsed JSON
    all_page_data = []

    for i, page in enumerate(pages):
        print(f"Sending page {i+1} to Ollama...")
        print(page)

        messages = [
            {"role": "template", "content": json.dumps(template)}
        ]

        # for ex in examples:
        #     messages.append({"role": "examples.input", "content": ex["input"]})
        #     messages.append({"role": "examples.output", "content": ex["output"]})

        messages.append({
            "role": "user",
            "content": f"Page {i+1}:\n{page}"
        })


        response = ollama.chat(
            model=model,
            messages=messages,
            options={"temperature": 0, "num_predict": 4096,"repeat_penalty": 1.1,"top_k": 1,"num_ctx": 8192}
        )
        print(response.message.content)

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
    for section in ['purchases', 'sales', 'interest']:
      data[section].sort(key=lambda x: (x.get('description') or '', x.get('date') or ''))
    for section in ['dividends']:
      data[section].sort(key=lambda x: (x.get('date') or ''))

    print(data)


# Create a new workbook and select the active sheet


    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Transactions"

    row = 1

    def write_row(cells, amount_cols=None, bold=False):
        """Write a row to the sheet.
        amount_cols: list of zero-based column indices to format as accounting
        """
        global row
        if amount_cols is None:
            amount_cols = []
        for col, value in enumerate(cells, start=1):
            cell = ws.cell(row=row, column=col, value=value)
            if col-1 in amount_cols and (
            isinstance(value, (int, float)) or (isinstance(value, str) and value.startswith("="))):
              cell.number_format = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
            if bold:
                cell.font = Font(bold=True)
        row += 1

    def add_space(rows):
        for i in range(rows): write_row([]) 

    # Purchases
    write_row(['Purchases:'])
    purchases_grouped = group_by_name(data['purchases'])
    for name, items in purchases_grouped.items():
        write_row([name],bold=True)
        add_space(1)
        write_row(['Date', 'Quantity', 'Amount'], bold=True)
        add_space(1)
        for item in items:
            # Amount is column 2 (0-based index)
            write_row([item['date'], f"{item['quantity']:.3f} shares" if item.get('quantity') else "", math.fabs(float(item['amount']))], amount_cols=[2])
        add_space(1)
        write_row(['','Total', 
                   f"=SUM({get_column_letter(3)}{row-len(items)-1}:{get_column_letter(3)}{row-1})", ], 
                   amount_cols=[2,3], bold=True)
        add_space(3)
    add_space(3)

    # Sales
    write_row(['Sales:'])
    sales_grouped = group_by_name(data['sales'])
    for name, items in sales_grouped.items():
        write_row([name], bold=True)
        add_space(1)
        write_row(['Date', 'Quantity', 'Carry Value', 'Sales Price', 'Gain', 'Loss'], bold=True)
        add_space(1)
        for item in items:
            gain_loss = item.get('realized_gain_loss', 0)
            carry_value = item.get('carry_value')
            if (carry_value is None or carry_value < .1 ) and gain_loss is not None:
                carry_value = item['amount'] - gain_loss
            # Columns 2=Amount, 3=Carry_Value, 4=Gain, 5=Loss
            write_row([
                item['date'],
                f"{item['quantity']:.3f} shares" if item.get('quantity') else "",
                float(carry_value) if carry_value is not None else None,
                float(item['amount']) if item.get('amount') else None,
                float(gain_loss) if gain_loss > 0 else 0,
                float(-gain_loss) if gain_loss < 0 else 0
            ], amount_cols=[2,3,4,5])
        add_space(1)
        write_row(['','Total', 
                   f"=SUM({get_column_letter(3)}{row-len(items)-1}:{get_column_letter(3)}{row-1})", 
                   f"=SUM({get_column_letter(4)}{row-len(items)-1}:{get_column_letter(4)}{row-1})", 
                   f"=SUM({get_column_letter(5)}{row-len(items)-1}:{get_column_letter(5)}{row-1})", 
                   f"=SUM({get_column_letter(6)}{row-len(items)-1}:{get_column_letter(6)}{row-1})"], 
                   amount_cols=[2,3,4,5,6], bold=True)
        add_space(3)
    

    # Dividends
    write_row(['Dividends:'])
    write_row(['Date', 'Description', 'Amount'], bold=True)
    for d in data['dividends']:
        write_row([d['date'], d['description'], float(d['amount'])], amount_cols=[2])
    add_space(1)
    write_row(['','Total',
                   f"=SUM({get_column_letter(3)}{row-len(data['dividends'])-1}:{get_column_letter(3)}{row-1})", ], 
                   amount_cols=[2,3], bold=True)
    add_space(3) 

    # Interest
    write_row(['Interest:'])
    write_row(['Date','Description', 'Amount'], bold=True)
    for i in data['interest']:
        write_row([i['date'], "Interest", float(i['amount'])], amount_cols=[3],bold=True)
    add_space(1)
    write_row(['','Total', 
                   f"=SUM({get_column_letter(3)}{row-len(data['interest'])-1}:{get_column_letter(3)}{row-1})", ], 
                   amount_cols=[2,3], bold=True)
    add_space(3) 

    # Save workbook
    wb.save('transactions.xlsx')