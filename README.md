# VBA of Wall Street

## Overview of Project

This project was undertaken to analyze the total daily volume and rates of return for various stocks.  

### Purpose

As part of a financial plan, many individuals invest in stocks from the stock market. When determing what stocks to invent in an individual may consider their personal ethics, in addition to, the rate of return for a particular stock. There are now more environmentally friendly or socially responsible stocks availabe for investors. However, an investor will still want to analyze the rate of return for these stocks to help diversify their portfolio. 

Our client is running analysis on green energy stocks to help his parents diversify their position in the market. We are working with stock data in Excel and using Visual Basic for Applications (VBA) to run the analysis. The analysis will help determine how the various stocks are performing over time. 

## Analysis

The analysis that was created allowed the user to see the total daily volume of stocks traded and the rate of return of the stocks. The analysis was completed for certain 'green' stocks and the rate of return was calculated by dividing the ending price by the starting price. Since we have access to two year's worth of data the user can select the year they want to analyze. The resulting data is formated so the user can quickly look at the table to find stocks that have a positive and negative rate of return before diving more deeply into the resulting analsysi. 

## Results

### Analysis Results

When conducting the stock analysis the conditional formatting for rate of return makes it obvious that there was a dramatic change between the rate of return for 2017 versus 2018. 

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)

The analysis for 2017, as shown above, demonstrates that all stocks, except one, experienced a positive rate of return. Looking closer at these numbers shows that four of the twelve stocks experienced a rate of return of over 100%. 

![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

The analysis for 2018 is quite different. As the above image shows, only two stocks had a positive rate of return with the other ten all having a negative rate of return. 

When doing a yearly comparison of the yearly data a couple of results appear. While the rate of return in 2018 for most of the stocks was negative, the Total Daily Volume for all stocks combined was higher in 2018 then in 2017.
 - Total Daily Volume for all stocks combined (2017): 3,166,639,100
 - Total Daily Volume for all stocks combined (2018): 3,306,038,200
 There are two stocks that had a positive rate of return in both 2017 and 2018: ENPH and RUN. When reviewing the differences it is important to note that there was a large decline in the stock market in 2018 that affected most stocks and this can be seen in our stock analysis. Conducting the analysis with 2019 and doing a comparison of these stocks over three years would help with the analysis of a specific stock is performing over time.

### Assignment Results

When conducting the analysis described above there was also an Assignment tasks that were also being completed. This was to analyze the differences between the original script and the refactored script. The refactored script was edited to make it more efficient to run. 

The run time for the original code is shown above.


The run time for the refactored script is shown below.


It can be seen that the refactored script runs much faster than the original script and this difference in speed would be a lot more evident with a larger data set. 

## Summary

- 1. What are the advantages or disadvantages of refactoring code? 

One advantage of refactoring code is that the makes the code more efficient. This can happen by remvoing extraneous lines of code or changing the parameters so items are not hard coded.  

A disadvantage of refactoring code is that is may result in changes to introduction of bugs. This could happen when parameters need to be changes for the factored code or due to developer inexperience.  

- 2. How do these pros and cons apply to refactoring the original VBA script? 

 - Pro: More efficient code
 For the second part of the assignment we used refactored code. The refactored code was more efficient. Adding the tickerIndex function increased the speed when the code was run and as noted above that speed difference was noticeable. The refactored code was also well documented meaning any developer could use the code and have a good understanding of what is happening in the code.

- Con: Introduction of bugs
The refactored code initially led to some bugs that I needed to figure out. Most of this due to my inexperience. It was necessary to review variables that were set and lines of code that were using other assumptions/parameters and determine if/how they needed to change for the refactored code. 
 
