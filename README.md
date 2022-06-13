# Green Stock Analysis

## Overview of Project

This project was initiated by a request from Steve, whose parents are interested in green investing. Steve just graduated with a degree in finance, and was concerned that his parents' stock portfolio was not diversified enough. They had put all their investments in DAQO New Energy Corporation, a company that makes silicon wafers for solar panels. Steve requested an analysis of twelve green energy stocks, including volume and rate of return. This analysis was completed in Excel using Visual Basic Application (VBA) to streamline the analysis process.

## Results

Full analysis of this stock data is available for download at [this link](https://github.com/ehalprin/stock-analysis/blob/main/VBA_Challenge.xlsm).

### Process of Initial Analysis

Using VBA, I wrote a code to analyze the twelve different stocks. I created a prompt for the user to enter what year they wanted to analyze. I created an array of all the tickers, and then created two different "for" loops to loop through the tickers and then all the rows, collecting information on volume, starting price, and ending price in the process. 

![Initial Analysis Code](https://github.com/ehalprin/stock-analysis/blob/main/Initial%202%20Loop%20Code.png))

Starting price and ending price were determined via code that identified when the current row was the first or last time a ticker appeared. From those two values, the yearly return could be determined. I then added conditional formatting to make it clear how each stock was performing. 

### Results of Initial Analysis

Here are the results:

![2017 Results](https://github.com/ehalprin/stock-analysis/blob/main/2017%20Stock%20Values.png)) ![2018 Results](https://github.com/ehalprin/stock-analysis/blob/main/2018%20Stock%20Values.png)

DQ stock performed well in 2017, with a 199.45% return. However, in 2018, DQ's return was -62.6%. ENPH and RUN were the only two stocks with positive returns in both 2017 and 2018, so those could make good additions to Steve's parents' stock portfolio. 

### Refactoring the Code

I then refactored the VBA code, so that it would run more efficiently if Steve wanted to analyze even more data. Instead of having two "for" loops with different variables, I used a new variable, "tickerIndex." This served as an index to four arrays: the array of tickers, and then three new arrays that I created for each ticker's volume, starting price, and ending price.

The tickerIndex started at 0, and, as each loop was completed, increased by one to move to the next ticker. The volumes of each ticker were added to the new tickerVolume array, matching the index of the ticker within the tickers array.

![Refactored Code](https://github.com/ehalprin/stock-analysis/blob/main/Refactored%20Code.png)

Because the code did not have nested for loops, the analysis ran more quickly and efficiently. Instead of running in 0.2 seconds, the analysis ran less than 0.1 second:

![VBA Challenge 2017](https://github.com/ehalprin/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA Challenge 2018](https://github.com/ehalprin/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring code obviously increases the speed at which it is processed. It removes unnecessary steps and streamlines the logic of the code. For extremely large data sets, increased speed and efficiency is key. One key disadvantage is that the process of refactoring could lead to mistakes or errors, adding new problems to code that is already working. 

### Refactoring this VBA Code

The refactored VBA code in this case worked much faster than the original, making it more useful in the future for processing even larger data sets. This seems to be largely due to having only one, rather than two nested, "for" loops. 

However, the process of refactoring was difficult; I had to make several attempts to understand theoretically what changes needed to be made, and then actually implement those changes. The first few refactored drafts led to erroneous data or code that simply wouldn't run as I tried to make the tickerIndex variable correspond correctly to the various arrays. However, ultimately it was worth the extra effort to create code that was more efficient and effective. 
