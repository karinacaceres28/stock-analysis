# stock-analysis
Performing analysis on Stock data to measure stock performance

# Stock Analysis using VBA

## Overview of Project
My friend Steve is a recent college graduate who obtained his degree in finance. Being his first clients, his parents are wanting to invest all their money into a single green energy stock. However, Steve believes the best option for his parents is for the funds to be more diversified. With data obtained for 12 stocks, I was able to generate outcomes for the total daily volume and the annual return for each stock using VBA in Excel. In doing so, this ensures the best investment for Steve’s parents.

### Purpose
While using VBA, the purpose of this project is to compare refactored and original solution code and determine the more efficient way to view an entire stock market over the last few years. In doing so we help Steve’s parents make the best investment for themselves.

## Results

### Original Code
The original code starts with formatting the output sheet on the “All Stocks Analysis” worksheet by activating the “All Stocks Analysis” sheet, putting the range for cell A1 value to the active year, and changing the headers for cells (3, 1-3) to say “Ticker,” “Total Daily Volume,” and “Return.” Next, we initialize an array for tickers 0 to 11. Following that I prepared for the analysis of tickers by initializing the starting price and ending price as doubles, activate the active year, and getting the number of rows to loop over. Then I moved on to looping through the tickers and remembering to reset the total volume to zero.  After, I looped through rows in the data by finding the total volume, starting price, and ending price for the current ticker. I output the data for the current ticker and added a button to put the selected year. Finally, adding some formatting to the cells. 

<img width="223" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110318652/191645242-1e30c85c-15ed-48ea-86eb-73c08aff1ccd.png">
