Ukraine Electricity Market Price Parser
This document describes, step by step, how the Python parser collects day-ahead electricity market (DAM) price data from Ukraine's Electricity Market Operator (OREE) and exports it to a formatted Excel file.
Source
Data is retrieved from: https://www.oree.com.ua/index.php/pricectr/data_view
How it works
1.	Set up the request.
The script defines the target URL, browser-like HTTP headers (User-Agent, Referer, X-Requested-With), and query parameters — the reporting month, market type (DAM), and price zone (IPS).
2.	Fetch the data.
A POST request is sent to the OREE endpoint. The server responds with JSON containing an HTML price table embedded in a "content" field.
3.	Parse the HTML table.
BeautifulSoup locates the <table> element inside the HTML fragment, and pandas reads it directly into a DataFrame (dates as rows, hours as columns).
4.	Reshape the data.
The wide table (Date | Hour 1 | Hour 2 | ...) is converted into a long format (Date | Hour | Price) using pandas melt(), which is easier to sort, filter, and load into other tools.
5.	Sort chronologically.
The Hour column is converted from text to numeric, then rows are sorted by Date and Hour so the output reads in the correct time order.
6.	Build the Excel file.
An openpyxl workbook is created. Column headers are written in row 1 with bold white text on a blue background, and the price data is written below starting in row 2.
7.	Auto-fit columns and save.
Each column's width is adjusted to fit its longest value, and the workbook is saved as an .xlsx file named after the query parameters (e.g. oree_prices_DAM_IPS_04_2026.xlsx).
Output
A single Excel file containing three columns — Date, Hour, and Price (грн/МВт.год) — with one row per hourly price observation for the selected month, market, and zone.
