# Stock Data Summary for All Quarters

## Overview
This VBA script is designed to summarize stock data across four quarterly sheets within an Excel workbook. It calculates the total quarterly change, total percentage change, and total stock volume for each ticker symbol. Additionally, it identifies the greatest percentage increase, greatest percentage decrease, and greatest total volume among all tickers.

## Features
- Checks for the existence of quarter sheets (Q1, Q2, Q3, Q4) before proceeding.
- Adds headers for summary columns in each quarter sheet.
- Creates a dictionary to store unique tickers and their summaries.
- Calculates the quarterly change and percent change for each ticker.
- Identifies tickers with the greatest increase, decrease, and volume.
- Outputs the summary to the specified columns in each quarter sheet.
- Auto-fits columns for better readability.

## How to Run
1. Open the Excel workbook that contains the stock data in separate sheets named Q1, Q2, Q3, and Q4.
2. Press `ALT + F11` to open the Visual Basic for Applications (VBA) editor.
3. Insert a new module and paste the provided VBA script into the module.
4. Close the VBA editor and run the script by pressing `F5` or by selecting 'Run Sub/UserForm' from the toolbar.

## Pre-requisites
- Ensure that each quarter sheet has stock data starting from the second row.
- The first row should contain headers, and the data should include ticker symbols, dates, and stock volumes.

## Output
The script will output the following in each quarter sheet:
- Ticker Symbol
- Total Quarterly Change
- Total Percentage Change
- Total Stock Volume
- Greatest % Increase (with corresponding ticker symbol)
- Greatest % Decrease (with corresponding ticker symbol)
- Greatest Total Volume (with corresponding ticker symbol)

## Notes
- The script assumes that the data is sorted by ticker symbol and date.
- If a quarter sheet does not exist, a message box will alert the user and the script will exit.
- Code source:  Learning assistant, google

## Disclaimer
This script is for educational purposes only. Please ensure that you have backed up your data before running the script as it will modify the contents of your workbook.



