# VBA of Wall Street

## Overview of Project
Steve wants some help analysing green stocks on Wall Street.  We were provided with 2 years of data for 12 different stocks to utilize.

### Purpose
The purpose of this project is to VBA in Excel to analyse green stocks to help make investment decisions.

## Analysis and Challenges

### Challenges and Difficulties Encountered
The difficulties that I had with this project were primarily setting up the tickerIndex variable with the output arrays.  It took me a few tries
to get the syntax correct and get them formatted properly.

## Results

- *Compare stock performance between 2017 and 2018*
Apparently 2017 was a much better year for green stocks on the stock market.  Only one of the stocks failed to increase it's value.  
![stock analysis for 2017](https://github.com/AndyHerron/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
In 2018 only 2 of the stocks were able to increase their values.  The other 10 stocks that we analyzed all had negative returns. 
![stock analysis for 2018](https://github.com/AndyHerron/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

- *Compare the execution times of the original script vs. the refactored script*
The refactored script was quicker for both years than the original script.  The original script did not use arrays for the various values and output the data after it analyzed each individual stock.  This required going back and forth between the worksheets after every loop.  Using arrays for the values allowed the refactored script to just go back to the output worksheet after everything had been calculated write the results with a new loop.

## Summary

- *What are the advantages or disadvantages of refactoring code?*
Refactoring code is intended to make code work more efficiently and be easier to read.  It will largely depend upon the intial code that you start with.

- *How do these pros and cons apply to refactoring the original VBA script?*
The original script did not use arrays for the various values and output the data after it analyzed each individual stock. This required going back and forth between the worksheets after every loop.  Using arrays for the values allowed the refactored script to just go back to the output worksheet after everything had been calculated write the results with a new loop.
