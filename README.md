# Stock Analysis

## Overview of Project
The purpose of this project was to refactor a Microsoft Excel VBA code to collect certain stock information in the year 2017 and 2018 and determine 
whether or not the stocks are worth investing. This process was originally completed in a similar format, however, the goal for this round was to 
increase the efficiency of the original code. 

## Results
The data includes two charts with stock information on 12 different stocks. The stock information contains a ticker value, the 
date the stock was issued, the starting, ending and adjusted ending price, the highest and lowest price, and the volume of the stock. The goal 
is to retrieve the ticker, the total daily volume, and the return on each stock.

Before refactoring the code, I copied the code that was needed to create the input box, chart headers, ticker array, and to activate the 
appropriate worksheet. The steps were then listed out in order to set the structure for the refactoring.

## Summary

Refactoring helps make our code cleaner and more organized. A few advantages of a cleaner code include design improvement and debugging. It may also 
benefit other users who view our projects because it becomes easier to read, as it is more concise and straightforward. 
The disadvantages may range from having applications that are too large to not having the proper test cases for the existing codes, which may ultimately 
pose some risk if we try to refactor our code. 

The biggest benefit that occurred as a result of the refactoring was an decrease in macro run time. 
The original analysis took approximately one 0.72, whereas our new analysis only took about approx 0.11 to run.
