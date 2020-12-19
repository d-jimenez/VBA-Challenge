# Stock Market Analysis

## Overview

The Stock Market Analysis project uses Excel Visual Basic for Applications to analyze yearly stock data information. The original data set contains daily stock information for a number of stocks, organized by year. The goal of the VBA script is to loop through the extensive data set, in order to build a summary table that calculates various yearly performance metrics for each stock. 

## Before You Begin

1. Pull the yearly data set and save as a macro enabled workbook (.xlsm).

2. Ensure that the developer tab within the excel ribbon has been enabled.

## Heads Up

1. The scrip is broken up into zzz major subroutines, each with a different funciton.

2. The zzz subroutine calls each of the individual subroutines and runs them as a single macro.

3. It is important that only the zzz subroutine is played to run, keeping in mind that by running only the zzz subroutine, the rest of the subroutines are called on. 

## Running the Code

1. Open Excel VBA.

2. Select Module 1.

3. Scroll to the subroutine located at the bottom of Module 1, named Sub zzz() and place cursor within subroutine. 

4. Select the play button within the VBA ribbon.

5. Allow the code to itterate through all of the data as well as each individual worksheet wothin the workbook. Keep in mind that this step may take a large amount of time due to the size of the data sample. 

## Results

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
   
Screenshots of the results for the first 20 unique stocks are included in the repository.
