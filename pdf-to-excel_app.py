#!/usr/bin/env python3
"""
pdf_to_excel.py

Extracts tabular data from a PDF purchase order, correctly merging rows
that span page breaks and avoiding duplicate EANs, then writes the result
to an Excel file.

Usage:
    python pdf_to_excel.py input.pdf -o output.xlsx

Dependencies:
    pip install pdfplumber pandas openpyxl
"""

import re
import argparse
import pdfplumber
import pandas as pd

EAN_REGEX = re.compile(r'(\d{8,13})')  # Adjust length if needed
NEW_ROW_REGEX = re.compile(r'^\s*(\d+)\s+')  # Lines starting with LP (number + space)

def extract_ean(text: str) -> str:
    """Find the first EAN (barcode) in the text."""
    m = EAN_REGEX.search(text)
    return m.group(1) if m else ''

def process_pdf(path: str) -> pd.DataFrame:
    """
    Reads the PDF, extracts all lines of text, and merges broken rows
    so that cells spanning pages stay together.
    """
    all_lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            # Split into individual lines
            lines = text.split('\n')
            all_lines.extend(lines)

    records = []
    current = None

    for line in all_lines:
        line = line.strip()
        if not line:
            continue

        # If line starts with an LP number => start of a new record
        if NEW_ROW_REGEX.match(line):
            # Save previous record
            if current:
                records.append(current)
            # Initialize new record dict
            # Split columns by two or more spaces (common delimiter in PDF text)
            cols = re.split(r'\s{2,}', line)
            lp = cols[0].strip()
            # You can expand this to parse other columns like 'index', 'name', 'vat', etc.
            description = ' '.join(cols[1:])  # everything else goes into description initially
            current = {
                'lp': lp,
                'description': description,
                'ean': '',
            }
        else:
            # Continuation of the previous record
            if not current:
                # If we see continuation without a current record, skip
                continue
            # If it's an EAN line, extract and assign
            if 'Kod kres' in line or 'EAN' in line:
                ean = extract_ean(line)
                if ean:
                    current['ean'] = ean
            else:
                # Otherwise it's a continuation of the description/name
                current['description'] += ' ' + line

    # Append the last record
    if current:
        records.append(current)

    # Build DataFrame
    df = pd.DataFrame(records, columns=['lp', 'description', 'ean'])

    # Remove exact duplicates of lp+ean (keep first)
    df = df.drop_duplicates(subset=['lp', 'ean'], keep='first')

    return df

def main():
    parser = argparse.ArgumentParser(
        description="Extract tables from PDF and write to Excel, merging "
                    "rows broken at page boundaries."
    )
    parser.add_argument('pdf', help="Path to input PDF file")
    parser.add_argument(
        '-o', '--output',
        default='output.xlsx',
        help="Path to output Excel file (default: output.xlsx)"
    )
    args = parser.parse_args()

    df = process_pdf(args.pdf)
    df.to_excel(args.output, index=False)
    print(f"✔ Saved {len(df)} rows to '{args.output}'")

if __name__ == '__main__':
    main()
