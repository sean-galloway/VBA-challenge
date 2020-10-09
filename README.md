# VBA-challenge
Week 2 of the Data Analysis Boot Camp

### Introduction
This project uses VBA to take stock ticker data from an existing spreadsheet, analyze the data, and summarizes the spreadsheet's results. The final step is saving the results to an output directory so that the original data remains untouched.

The analysis performed is:
* Calculate the yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* Calculate the percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* Sum the total stock volume of the stock.
After completing the analysis, the VBA code changes the "Yearly Change" cells background color:
* Values greater than or equal to zero are colored green
* Values less than zero are colored red

Further anlysis that is performed is:
* Determine the greatest percent increase
* Determine the greatest percent decrease
* Determine the greatest total volume

### Usage
Open the master VBA macro file, [VBA-challenge.xlsm](VBA-challenge.xlsm). Within this file, there are three worksheets:
* Lookup: this sheet holds the file names and the legal worksheets for the resources data. This data gets populated in the drop-down menus on the other worksheets.
* SingleSheet: this sheet has two drop-down menus (these select the workbook and worksheet to analyze) and a "Run the Single Script" button. The VBA code will only analyze the data on a single sheet. SingleSheet is a code testing sheet.
* MultiSheet: this sheet has one drop-down menu (to select the workbook to analyze) and a "Run the Multi Script" button. The VBA code analyzes all of the sheets in the workbook. The multi-script is a production-quality version of the VBA code.

If the user chooses the SingleSheet option, select the workbook and the worksheet by clicking blue cells. Upon choosing the workbook and worksheet, click on the "Run the Single Script" button. The VBA code opens the workbook, runs the analysis on the one worksheet, then saves it to the Output directory.

If the user chooses the MultiSheet option, select the workbook by clicking the blue cell. Upon choosing the workbook, click on the "Run the Multi Script" button. The VBA code opens the workbook, runs the analysis on all worksheets, then saves it to the Output directory.

### File List
Below is a list of the directories used on this project:
* Images- this contains all the required images for this project. These consist of screenshots of the data generated for the Multiple_year_stock_data.xlsx file analysis.
* Output- this contains the output from running the VBA script.
* Resources- this contains the raw reference data.
* VBA_Scripts- this contains the three VBA script files

