# All Stocks Analysis

## Overview of Project

This project is an analysis of All Stock Tickers for the year 2017 and 2018 using VBA and excel. The exercises in this project were to analyze a defined list of stocks using VBA (particularly `For` loops and `If`/`Then` statements) to iterate over the data to calculate Daily Volume and Returns. The last part of the exercise was to include a code that computes how quickly the code was executed.


## Results

### Analysis for All Stock 2017

For the year 2017, All tickers had positive returns except for TERP which had a total of $139,402,800 Total Daily Volume and a -7.2% Return. Ticker SPWR had the highest Total Daily Volume in terms of dollar amount. However, Ticker DQ had the highest total return with a percentage of 199.4% at $35,796,200 Total Daily Volume.


![All_Stock_2017](https://raw.githubusercontent.com/Mishabatoon/stock-analysis/main/Resources/VBA_Challenge_2017.png)


### Analysis for All Stock 2018

For the year 2018, Even though Ticker DQ had tripled its Total Daily Volume since last year, it had the lowest return percentage for this year with a total return of -62.6%. Out of (12) twelve Tickers, only (2) two had a positive percentage return for 2018. Ticker RUN had the highest return with a positive percentage of 84%.

![All_Stock_2018](https://raw.githubusercontent.com/Mishabatoon/stock-analysis/main/Resources/VBA_Challenge_2018.png)

### Execution Times of the Original VS Refactored Script

The execution time for running the code for 2017 took 0.29 seconds, while the code for 2018 ran in 0.31 seconds. Both are fast and efficient.

Personally, I like to create my own VBA when analyzing (1) one set of data because I can set all the parameters and structure of the script. However, when analyzing multiple sets of data, refactoring is a lot easier and faster. When refactoring, I can just make minor changes and introduce new parameters without making a major modification to the VBA script.


## Summary

### Advantages of Refactoring Code

- Ultimately saves time and effort to write the code.
- Loops and If statements are already set.
- When refactoring, we only make minor changes like adding and removing parameters, restructuring the code without totally modifying the external behavior.
### Disadvantages of Refactoring Code
- When making minor changes to the code, it can introduce new bugs and errors.
- Changing variables can cause a major problem not just with the structure of the code but also to the results of the data.

All these Pros and Cons should be taken into consideration when creating a VBA script. Refactoring takes practice and discipline and must be fully undestood in order to maintain the integrity of the script.



