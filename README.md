# VBA-challenge
VBA script to analyze real stock market data

## Background

This is a VBA script that will analyze the following daily stock market data for a given year's worth of transactions.  Each tab in the workbook may have a years worth of daily transaction data by ticker:
  1. Ticker
  2. Date
  3. Open
  4. High
  5. Low
  6. Close
  7. Volume
  
Each tab in the data file includes a year's worth of stock market data broken down to daily detail. Based on the fields above, the script will loop through all the stocks in a year for each worksheet and output the following data consolidated by ticker symbol.
    1. Ticker symbol
    2. The yearly change (end of year closing price - beginning of year opening price).
    3. The percent change in stock price for the year.
    4. The total trading volume of each stock for the year.
    5. Any stocks that increased in price that year are highlighted in green. Any stocks that decreased in price in a given year highlighted in red.
    
## Stock Selection
In addition to compiling the data for each ticker symbol the VBA script also selects the following ticker symbols from each year with the following criteria:

  1. Stock with the greatest % increase in price
  2. Stock with the greatest % decrease in price
  3. Stock with the greatest total trading volume

## Link to .vbs files containing script

  1. Multiple_year_stock_data_script:
  https://github.com/jbski/VBA-challenge/blob/master/Multiple_year_stock_data.vbs  
  
  
  2. alphabetical_testing_data_script:
  https://github.com/jbski/VBA-challenge/blob/master/alphabetical_testing.vbs
  
  
## Link to .xlsm file containing sample data
A file including sample data has been included at the following link.  The file containing the full set of multiple year stock data 
is ~115MB and exceeds the 25MB Github limit.

Sample Stock Data:
https://github.com/jbski/VBA-challenge/blob/master/alphabetical_testing.xlsm
    
    
## Screen shots of results:

2014 Stock Results:
https://github.com/jbski/VBA-challenge/blob/master/2014%20Stock%20Results.PNG

2015 Stock Results:
https://github.com/jbski/VBA-challenge/blob/master/2015%20Stock%20Results.PNG

2016 Stock Results:
https://github.com/jbski/VBA-challenge/blob/master/2016%20Stock%20Results.PNG

    
