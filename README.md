# Stock Market Analysis

## Overview

The Stock Market Analysis project uses Excel Visual Basic for Applications to analyze yearly stock data information. The original data set contains daily stock information for a number of stocks, organized by year. The goal of the VBA script is to loop through the extensive data set, in order to build a summary table that calculates various yearly performance metrics for each stock. 

## Before You Begin

1. Pull the yearly data set and save as a macro enabled workbook (.xlsm).

2. Ensure that the developer tab within the excel ribbon has been enabled.

## Heads Up

1. The scrip is broken up into four major subroutines, each with a different funciton.

2. The MultiSheetFullProcess() subroutine calls each of the individual subroutines and runs them as a single macro.

3. It is important that only the MultiSheetFullProcess() subroutine is played to run, keeping in mind that by running only the zzz subroutine, the rest of the subroutines are called on. 

## Running the Code

1. Open the Multiple_year_stock_data.xlsm file and import the VBA script.

2. Select Module 1.

3. Scroll to the subroutine located at the bottom of Module 1, named Sub MultiSheetFullProcess() and place cursor within subroutine. 

4. Select the play button within the VBA ribbon.

5. Allow the code to itterate through all of the data as well as each individual worksheet wothin the workbook. After the code is finished running a messagebox will appear stating the ampunt of time it took to run the code.

## Output

The VBA script will ouput a summary table containing the following information for each unique stock symbol:

1. The Yearly Change in stock price (Yearly Close Price minus Open Price).
2. The Percent Change in stock price (Yearly Change divided by Open Price).
        - Percent Change is formated so than any stock that had a positive percent change is highlighted in green, while those with negative changes are in red.
3. The Total Stock Volume traded, (Sum of Stock Volumes).
    
The bonus material includes: 
   
1. The Maximum Percent Change and its associated stock symbol.
   
2. The Minimum Percent Change and its associated stock symbol. 
   
3. The Largest Stock Volume and its associated stock symbol. 
   
4. Thescript also itterates through of of the yearly worksheets within the workbook.
    
## Brief Results Overview
   
1. Bank of America Corp. (BAC) had the largest total stock volume in 2014, 2015 and 2016 with an average stock volume of 23,433,922,066.6667 yearly.

2. The maximum yearly percent stock change  for the 2014, 2015 and 2016 data was for SandRidgeEnergy of 11,675% increasing from $0.20 to $23.75 for the 2016 year.

3. Kinder Morgan Inc. fell by 98.49 percent in 2015, recording the largest percent decrease in stock prices for all stocks in the 2014, 2015 and 2016 data.
   
* Screenshots of the results for the first 20 unique stocks for the years are included in the repository.

