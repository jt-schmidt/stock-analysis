# stock-analysis
Module_2_Using_VBA_with_Excel

## Background
Purpose of this assignment was to build skill with Visual Basic by performing an analysis of various stocks.  
For the data set:
* There were 12 stock symbols
* Data was separated by year into separate worksheets
* Within each sheet, data was organized in table alphabetically by stock symbol

## [Original Code](VBA_original.vbs)
This is the code used during the lesson to perform an initial analysis of the stock data.  Skills explored include:
* For loops
* If Then statements
* Variable definition and type including arrays
* Text and cell formatting
* Basic math and logical operators

## [Refactored Code](VBA_challenge.vbs)
This is the code used during independent challenge.  Primary goal was to introduce a *tickerIndex* variable which could be used as an address to tie a specific stock symbol with the starting price, ending price, and calculated total volume for that same stock symbol.

## Conclusions
This was a great exploration of multiple facets of Visual Basic language and the strength it can provide in performing repetitive analysis at the click of a button.  In addition to introducing to basic programming concepts it also taught a lesson in refactoring or editing code to improve performance.

### Refactor in General
Refactoring code provides a couple of key advantages:
1. Improving efficiency when running.  This reduces strain on system resources & reduces wait time for end user.
2. Introducing & inceasing exception handling.  Initial iteration of creating code often overlooks error handling.
    Example:  Return Calculation may run into divide by zero error if [tickerStartingPrices is zero](https://blog.mywallst.com/can-stocks-go-to-zero/)
```
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
[The primary disadvantage to refactoring code is the ROI.](https://thenewstack.io/refactoring-is-not-bad-until-it-is/)  Before refactoring, the issue being addressed needs to be weighed against the needs of the business, the impact the code is having on the business, and the resources available to refactor.

### Refactor with respect to Stock Analysis
Refactor specific to this Stock Analysis was as a learning exercise--so the ROI is demonstrated.  
The primary advantage was to conceptually teach managing a multi-dimensional array through a series of loops. 

Stock Symbol | Starting Price | Ending Price | Total Volume
------------ | -------------  | ------------- | -------------  |
AY | 19.47 | 21.21 | 136070900

This is a huge advantage in programming and data analysis.  
The primary disadvantage of this specific refactor?  This is difficult to say & is subject to opinion.  
* Perhaps the *tickerIndex* could have been used to step through the stock data and exit the loop when the symbol changed?  
* What if the stock data was not arranged alphabetically by stock symbol?  The code could be made to search through the data and identify startPrice based on the Date field.
There are numerous possibilities.
