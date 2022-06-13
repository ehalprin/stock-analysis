# Green Stock Analysis

## Overview of Project

This project was initiated by a request from Steve, whose parents are interested in green investing. Steve, who just graduated with a degree in finance, was concerned that their stock portfolio was not diversified enough, as they were fully invested in DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. Steve requested an analysis of DAQO stock, including volume and rate of return, as well as that information on a few other green energy stocks. This analysis was completed in Excel using Visual Basic Application (VBA) to streamline the analysis process.

## Results

Full analysis of this stock data is available for download at [this link](https://github.com/ehalprin/stock-analysis/blob/main/green_stocks.xlsm).

Using VBA, I wrote a code to analyze twelve different stocks. I created a prompt for the user to enter what year they wanted to analyze. I created an array of all the tickers, and then created two different "for" loops to loop through the tickers and then all the rows, collecting information on volume, starting price, and ending price in the process. 

![Initial Analysis Code](https://github.com/ehalprin/stock-analysis/blob/main/Initial%202%20Loop%20Code.png))

Starting Price and Ending price were determined via code that identified when the current row was the first or last time a ticker appeared. From those two values, the yearly return could be determined. I then added conditional formatting to make it clear how each stock was performing. Here are the results:

![2017 Results](https://github.com/ehalprin/stock-analysis/blob/main/2017%20Stock%20Values.png)) ![2018 Results](https://github.com/ehalprin/stock-analysis/blob/main/2018%20Stock%20Values.png)



## Summary
