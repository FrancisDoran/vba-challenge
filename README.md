# Stock Analysis Tool

## Overview
This Stock Analysis Tool is an Excel-based VBA (Visual Basic for Applications) application designed to analyze stock data across multiple worksheets in an Excel workbook. It calculates yearly changes, percent changes, and total stock volumes for each stock ticker. Additionally, it identifies the greatest percent increase, greatest percent decrease, and greatest total volume for stocks.

## Features
- **Yearly Change Calculation**: Computes the yearly change in stock price for each ticker.
- **Percent Change Calculation**: Calculates the percent change in stock price for each ticker.
- **Total Stock Volume Calculation**: Aggregates the total stock volume for each ticker.
- **Highlighting Stock Performance**: Highlights positive yearly changes in green and negative changes in red.
- **Analysis Across Multiple Sheets**: Performs stock analysis on every worksheet within the workbook.
- **Summary Metrics**: Identifies and displays the tickers with the greatest percent increase, greatest percent decrease, and greatest total volume.

## How to Use
1. **Prepare Your Data**: Ensure each worksheet in your Excel workbook contains stock data with tickers, start prices, end prices, and volumes.
2. **Run the Tool**: Execute the `challengeAllSheets` macro. This macro processes each sheet in turn.
3. **View Results**: After running the macro, check each worksheet for a summary table with calculated metrics.
