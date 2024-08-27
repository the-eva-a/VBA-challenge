# VBA Stock Market Analysis 

## Description
This project showcases how one can use VBA scripting to automate the analysis of stock market data from multiple financial quarters. The script handles stock data, calculating key financial measurements and applying visual indicators with conditional formatting.

## Table of Contents
- [Description](#description)
- [Overview](#overview)
- [Installation](#installation)
- [Output](#output)
- [Possible Future Improvements](#possible-future-improvements)
- [Acknowledgements](#acknowledgements)

## Overview
The VBA Stock Market Data Analysis project is a tool designed to automate the analysis of stock data across multiple quarters using VBA scripting. Data is stored with financial quarters stored on individual worksheets in a single workbook. The script analyzes the stock data, calculates the percent change and total volume for all stocks as well as the overall percent increase, percent decrease, and greatest overall stock volume. 

## Installation

To set up this project locally:

1. **Clone the repository:**
    ``` bash
    git clone https://github.com/the-eva-a/VBA-challenge.git
    cd VBA-challenge
    ```
2. **Open the Excel Workbook**
    - Locate and open the Excel file named `Multiple_year_stock_data.xlsm` included in the repository.

3. **Import VBA Scripts**
    - Open the Visual Basic for Applications (VBA) editor in Excel by pressing `Alt` + `F11`.
    - Import the scripts:
         - `ClearAllFormatting.bas`: This script removes all existing formatting from the workbook to ensure a clean starting point.
         - `StockInfo.bas` : This is the main script that performs the stock data analysis and applies appropriate formatting.
    - To import:
        - In the VBA editor, go to `File > Import File...`.
        - Select the `.bas` files one by one and click Open.

4. **Run the Scripts**
    - In the VBA editor, ensure that both scripts are properly imported under the correct workbook.
    - Execute the scripts in the following order:
        1. **ClearAllFormatting.bas**
            - This will clear any existing formatting from your dataset.
        2. **StockInfo.bas**
            - This will perform the analysis, calculate all necessary metrics, and apply conditional formatting to the results.
    - To run a script:
        - Select the desired script/module.
        - Press `F5` or go to `Run > Run Sub/UserForm`.

5. **View the Results**
    - Return to the Excel workbook to view the analyzed data.
    - The results will include:
        - Ticker symbols.
        - Quarterly and percentage changes in stock prices.
        - Total stock volume.
        - Highlights for the greatest percentage increase, decrease, and highest total volume.
        - Conditional formatting with green indicating positive changes and red indicating negative changes.

**Note:** Ensure that macros are enabled in Excel to allow the VBA scripts to run correctly.

## Output

The script outputs the following data for each stock:
- **Ticker Symbol:** Identifies each stock.
- **Quarterly Change:** The difference between the opening price at the beginning and the closing price at the end of the quarter.
- **Percentage Change:** The percentage change in price over the quarter.
- **Total Stock Volume:** The sum of the trading volume for the quarter.
- **Greatest Metrics:** Identifies stocks with the greatest percentage increase, decrease, and total volume.
 
![Excel worksheet displaying the results after running the StockInfo VBA script. The sheet includes columns for stock tickers, dates, opening prices, high and low prices, closing prices, volumes, quarterly changes, percentage changes, and total stock volumes. Positive percentage changes are highlighted in green, and negative changes in red. The right side of the sheet shows the ticker symbols for the stocks with the greatest percentage increase, decrease, and total volume, along with their respective values](/screenshots/StockInfoOutput.png)
*Figure 1: Excel worksheet displaying the results after running the `StockInfo` VBA script.* 

![Excel worksheet displaying raw stock market data. The sheet includes columns for stock tickers, dates, opening prices, high and low prices, closing prices, and volumes. The data is unformatted, showing the clean state of the worksheet before further analysis. This is the expected format for data before running the StockInfo VBA script for analysis and conditional formatting. Data can be brought back to this form after using the StockInfo script by using the ClearAllFormatting script.](/screenshots/StockInfoInput.png)
*Figure 2: Excel worksheet displaying the proper data format and the results of the `ClearAllFormatting` VBA script.*

## Possible Future Improvements

If I were to continue my learning with VBA and this project, some possible future improvements could be:
- Exploring more complex financial metrics like price-to-earnings ratio.
- Adding comparisons between a specific stock from quarter to quarter.
- Creating a more user-friendly experience and menu so that those who are less comfortable with VBA can also use the program.

## Acknowledgements

This project is part of the curriculum in the edX Data Analytics Bootcamp. Special thanks to my instructors and peers for their support and guidance.
