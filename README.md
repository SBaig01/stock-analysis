# Stock-Analysis

## Overview of Project

### Purpose
Steve is looking to expand the dataset for analysis to include the entire stock market over the last few years, and not just 12 company stock tickers for 2 years. To ensure that our code runs efficiently this analysis was conducted to determing if refactoring the code would make the VBA script run faster. 

## Results

### 2017 Results
Before refactoring the code, the script took 0.80 seconds to execute. After refactoring, it took 0.11 seconds.

![2017_Results](https://github.com/SBaig01/stock-analysis/blob/b5e877e486e094335a1b7b730b87511c28617d5d/VBA_Challenge_2017.png)

### 2018 Results
Before refactoring the code, the script took 0.79 seconds to execute. After refactoring, it took 0.12 seconds.

![2018_Results](https://github.com/SBaig01/stock-analysis/blob/b5e877e486e094335a1b7b730b87511c28617d5d/VBA_Challenge_2018.png)

## Summary

### What are the advantages or disadvantages of refactoring code?
- Advantages: The script runs faster and saves time, the code is easier to modify and is easier to read as it is less cluttered.
- Disadvantages: It takes time to refactor code and that means increased costs for business. If you refactor a working code, you may damage it.

### How do these pros and cons apply to refactoring the original VBA script?
While the original VBA script worked and returned the desired values, the running time was almost 8 times more than the refactored script. By introducing a tickerIndex we were able to call upon the appropriate 3 Arrays of Volumes, Starting Prices and Ending Prices for the All Tickers array and consecutively increase the tickerIndex. This made the script run faster as the nested For Loop would generate the output Arrays without having to start at the beginning of the initial For Loop to find the appropriate Ticker.
