# 01-Electricity-Market-Pasrser-Ukraine
This repository is dedicated to a Python parser which collects price data from electricity market operator in Ukraine.

This documents describes the code in Python which collects price series from the Ukraine's Electricity Market Operator.
The code processes data from the web page https://www.oree.com.ua/index.php/pricectr/data_view.
# ============================================================
# IMPORTS
# ============================================================
import requests                                      # send HTTP requests to fetch data from the website
from bs4 import BeautifulSoup                         # parse HTML content and extract the price table
from io import StringIO                               # wrap HTML string so pandas can read it as a "file"
import pandas as pd                                   # reshape/sort tabular data
from openpyxl import Workbook                         # build the output .xlsx file
from openpyxl.styles import Font, PatternFill, Alignment  # style Excel headers (bold, colored, centered)

# ============================================================
# REQUEST SETUP
# ============================================================
# Endpoint that returns the day-ahead market (DAM) price table as JSON
url = "https://www.oree.com.ua/index.php/pricectr/data_view"

# Headers that mimic a real browser request. Without these, the server
# may reject the request or return a different (non-AJAX) response.
headers = {
    "User-Agent": "Mozilla/5.0",                       # identify as a standard browser
    "Referer": "https://www.oree.com.ua/index.php/pricectr",  # page the request "comes from"
    "X-Requested-With": "XMLHttpRequest",               # tells server this is an AJAX/API call
}

# Parameters sent to the server: which month, which market, which price zone
payload = {"date": "04.2026", "market": "DAM", "zone": "IPS"}

# Build a descriptive output filename based on the query parameters,
# e.g. "oree_prices_DAM_IPS_04_2026.xlsx"
filename = f"oree_prices_{payload['market']}_{payload['zone']}_{payload['date'].replace('.', '_')}.xlsx"

# ============================================================
# FETCH DATA
# ============================================================
resp = requests.post(url, data=payload, headers=headers)  # POST the query and get the response
resp.encoding = "utf-8"                                    # ensure Cyrillic text decodes correctly

# The API returns JSON with an HTML table embedded inside a "content" field
content = resp.json().get("content", "")

# ============================================================
# PARSE HTML TABLE INTO A DATAFRAME
# ============================================================
soup = BeautifulSoup(content, "html.parser")   # parse the HTML fragment
table = soup.find("table")                     # locate the <table> element with the price data

# pandas can read an HTML table directly, but needs a file-like object,
# so StringIO wraps the table's HTML string in memory
df = pd.read_html(StringIO(str(table)))[0]     # [0] takes the first (only) table found

# ============================================================
# RESHAPE: WIDE (Date | Hour1 | Hour2 | ...) -> LONG (Date | Hour | Price)
# ============================================================
# The original table has one row per date and one column per hour.
# melt() converts it to one row per (date, hour) pair, which is easier
# to sort, filter, and load into other tools.
df_long = df.melt(
    id_vars=df.columns[0],                     # keep the Date column fixed
    var_name="Hour",                           # former column headers (hours) become a "Hour" column
    value_name="Price (грн/МВт.год)"           # former cell values become a "Price" column
)
df_long.rename(columns={df.columns[0]: "Date"}, inplace=True)  # rename the first column to "Date"
df_long = df_long[["Date", "Hour", "Price (грн/МВт.год)"]]      # reorder columns for clarity

# ============================================================
# SORT DATA CHRONOLOGICALLY
# ============================================================
# "Hour" was extracted from column headers, so it's currently text (e.g. "1", "2", ...);
# convert to numeric so sorting is 1, 2, 3... instead of alphabetical (1, 10, 11, 2...)
df_long["Hour"] = pd.to_numeric(df_long["Hour"], errors="coerce")
df_long = df_long.sort_values(["Date", "Hour"]).reset_index(drop=True)

# ============================================================
# BUILD THE EXCEL WORKBOOK
# ============================================================
wb = Workbook()
ws = wb.active
ws.title = "Prices"

# Header cell styling: white bold text on a blue background, centered
header_font = Font(bold=True, name="Arial", color="FFFFFF")
header_fill = PatternFill("solid", start_color="4472C4", end_color="4472C4")
header_align = Alignment(horizontal="center")

# Write column headers in row 1 and apply the styling defined above
for col_idx, col_name in enumerate(df_long.columns, start=1):
    cell = ws.cell(row=1, column=col_idx, value=str(col_name))
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align

# Write the data rows starting at row 2 (row 1 is the header)
for row_idx, row in df_long.iterrows():
    for col_idx, value in enumerate(row, start=1):
        ws.cell(row=row_idx + 2, column=col_idx, value=value)

# ============================================================
# AUTO-FIT COLUMN WIDTHS
# ============================================================
# Set each column's width based on its longest cell value,
# so text isn't cut off when the file is opened
for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = max_len + 4

# ============================================================
# SAVE AND DOWNLOAD (Google Colab only)
# ============================================================
wb.save(filename)              # write the .xlsx file to disk

from google.colab import files  # Colab-specific helper to trigger a browser download
files.download(filename)        # prompt the user to download the generated file

