# cleanup_and_build_excel.py
"""
Usage: python cleanup_and_build_excel.py [input.xlsx] [output.xlsx]

This script reads all sheets from the input Excel file, writes cleaned versions
and a basic Report_Print sheet into the output file. It also converts Gregorian
Excel dates to Jalali (Persian) date strings for detected date columns.

Requirements:
  pip install pandas openpyxl jdatetime numpy

Notes:
 - This script runs locally. It expects the input file to be available on your
   machine (e.g., after you clone the repo). It cannot modify binary Excel
   files inside the runner environment here.
 - The script attempts to detect date columns by column name keywords ('date',
   'تاریخ', 'زمان') and by dtype. It will convert datetimes to Jalali strings
   (YYYY/MM/DD) in the cleaned sheets and the report summary.

What it produces:
 - "سیستم ثبت برگه ورود خروج نهایی - fixed.xlsx" in the same folder (or the
   provided output path). Contains Data_Raw_* sheets and Data_Clean_* sheets,
   plus Report_Print and README sheet placeholders.

"""
import sys
import os
from datetime import datetime
import pandas as pd
import numpy as np
import jdatetime


def to_jalali_str(dt):
    if pd.isna(dt):
        return ""
    if isinstance(dt, (pd.Timestamp, datetime)):
        g = dt.to_pydatetime() if isinstance(dt, pd.Timestamp) else dt
        j = jdatetime.datetime.fromgregorian(datetime=g)
        return f"{j.year:04d}/{j.month:02d}/{j.day:02d}"
    # Try parse string
    try:
        parsed = pd.to_datetime(dt, dayfirst=True, errors='coerce')
        if pd.isna(parsed):
            return str(dt)
        j = jdatetime.datetime.fromgregorian(datetime=parsed.to_pydatetime())
        return f"{j.year:04d}/{j.month:02d}/{j.day:02d}"
    except Exception:
        return str(dt)


def detect_date_columns(df):
    date_cols = []
    keywords = ['date', 'تاریخ', 'time', 'زمان']
    for col in df.columns:
        c = str(col).lower()
        if any(k in c for k in keywords):
            date_cols.append(col)
            continue
        # dtype check
        if np.issubdtype(df[col].dtype, np.datetime64):
            date_cols.append(col)
    return date_cols


def clean_df(df):
    # Drop fully empty rows/columns
    df = df.copy()
    df.dropna(axis=0, how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    # If first row looks like header (no repeated values), use it
    # Ensure columns are strings
    df.columns = [str(c).strip() for c in df.columns]

    # Trim string columns
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip().replace({'nan': np.nan})

    # Convert numeric-like
    for col in df.columns:
        # try to coerce numeric
        try:
            df[col] = pd.to_numeric(df[col], errors='ignore')
        except Exception:
            pass

    # Detect date columns and convert to datetime if possible
    date_cols = detect_date_columns(df)
    for col in date_cols:
        try:
            df[col] = pd.to_datetime(df[col], errors='coerce')
        except Exception:
            pass

    return df, date_cols


def main():
    if len(sys.argv) > 1:
        input_path = sys.argv[1]
    else:
        # default to the filename in the repo
        input_path = '‎⁨سیستم ثبت برگه ورود خروج نهایی⁩.xlsx'

    output_path = sys.argv[2] if len(sys.argv) > 2 else 'سیستم ثبت برگه ورود خروج نهایی - fixed.xlsx'

    if not os.path.exists(input_path):
        print(f"Input file not found: {input_path}")
        print("Please place the original .xlsx file in the same folder as this script or pass the path as the first argument.")
        sys.exit(1)

    print(f"Loading workbook: {input_path}")
    sheets = pd.read_excel(input_path, sheet_name=None, engine='openpyxl')

    writer = pd.ExcelWriter(output_path, engine='openpyxl')

    summary = []

    for name, df in sheets.items():
        print(f"Processing sheet: {name} (shape={df.shape})")
        # Write raw copy
        raw_name = f"Data_Raw_{name}"
        df.to_excel(writer, sheet_name=raw_name, index=False)

        # Clean
        cleaned, date_cols = clean_df(df)

        # Convert detected date columns to Jalali strings
        for col in date_cols:
            cleaned[col] = cleaned[col].apply(to_jalali_str)

        clean_name = f"Data_Clean_{name}"
        cleaned.to_excel(writer, sheet_name=clean_name, index=False)

        # build a small summary
        total_rows = len(cleaned)
        sample_dates = []
        for col in date_cols:
            # find min/max from original parsed datetimes
            try:
                orig = pd.to_datetime(df[col], errors='coerce')
                if orig.dropna().shape[0] > 0:
                    mn = orig.min()
                    mx = orig.max()
                    sample_dates.append((col, to_jalali_str(mn), to_jalali_str(mx)))
            except Exception:
                pass

        summary.append({'sheet': name, 'rows': total_rows, 'date_columns': date_cols, 'date_sample': sample_dates})

    # Create Report_Print sheet
    from openpyxl import Workbook
    # We already have writer.openpyxl workbook; write a simple report sheet at the end
    # Create a small DataFrame for the report summary
    rep_rows = []
    for s in summary:
        rep_rows.append({
            'Sheet': s['sheet'],
            'TotalRows': s['rows'],
            'DateCols': ','.join(s['date_columns']) if s['date_columns'] else ''
        })
    rep_df = pd.DataFrame(rep_rows)
    rep_df.to_excel(writer, sheet_name='Report_Print', index=False)

    # Create README sheet content
    readme_lines = [
        'README for cleaned file',
        '',
        'This workbook was generated by cleanup_and_build_excel.py',
        'Sheets included:',
    ]
    for s in summary:
        readme_lines.append(f" - Original sheet: {s['sheet']} -> Data_Raw_{s['sheet']}, Data_Clean_{s['sheet']}")
    readme_lines.append('')
    readme_lines.append('Report_Print contains a short summary per sheet.')
    readme_df = pd.DataFrame({'Note': readme_lines})
    readme_df.to_excel(writer, sheet_name='README', index=False)

    print(f"Saving cleaned workbook to: {output_path}")
    writer.save()
    print('Done.')


if __name__ == '__main__':
    main()
