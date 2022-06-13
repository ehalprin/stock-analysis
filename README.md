# Green Stock Analysis

## Overview of Project

This project was initiated by a request from Steve, whose parents are interested in green investing. Steve, who just graduated with a degree in finance, was concerned that their stock portfolio was not diversified enough, as they were fully invested in DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. Steve requested an analysis of DAQO stock, including volume and rate of return, as well as that information on a few other green energy stocks. This analysis was completed in Excel using Visual Basic Application (VBA) to streamline the analysis process.

## Results

Full analysis of this stock data is available for download at [this link](https://github.com/ehalprin/stock-analysis/blob/main/green_stocks.xlsm).

Using VBA, I wrote a code to analyze twelve different stocks. I created a prompt for the user to enter what year they wanted to analyze. I created an array of all the tickers, and then created two different "for" loops to loop through the tickers and then all the rows, collecting information on volume, starting price, and ending price in the process. 

![Initial Analysis Code](https://github.com/ehalprin/stock-analysis/blob/main/Initial%202%20Loop%20Code.png))

Starting Price and Ending price were determined via code that identified when the current row was the first or last time a ticker appeared. From those two values, the yearly return could be determined. I then added conditional formatting to make it clear how each stock was performing. Here are the results:

![2017 Results](https://github.com/ehalprin/stock-analysis/blob/main/2017%20Stock%20Values.png)) ![2018 Results](https://github.com/ehalprin/stock-analysis/blob/main/2018%20Stock%20Values.png)

DQ stock performed well in 2017, with a 199.45% return. However, in 2018, DQ's return was -62.6%. ENPH and RUN were the only two stocks with positive returns in both 2017 and 2018, so those would make good additions to Steve's parents' stock portfolio. 

I then refactored the VBA code, so that it would run more efficiently if Steve wanted to analyze even more data. Instead of having two "for" loops with different variables, I used a new variable, "tickerIndex." This served as an index to four arrays: the array of tickers, and then three new arrays that I created for each ticker's volume, starting price, and ending price.

The tickerIndex started as 0, and, as each loop was completed, increased by one to move to the next ticker. The volumes of each ticker were added to the new tickerVolume array, matching the index of the ticker within the tickers array.

![Refactored Code](https://github.com/ehalprin/stock-analysis/blob/main/Refactored%20Code.png)

Because the code did not have nested for loops, the analysis ran more quickly and efficiently. Instead of running in 0.2 seconds, the analysis ran less than 0.1 seconds:

![VBA Challenge 2017](https://github.com/ehalprin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA Challenge 2018](https://github.com/ehalprin/stock-analysis/blob/main/VBA_Challenge_2018.png)


## Summary
