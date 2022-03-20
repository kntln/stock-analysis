# Stock Analysis With Excel VBA

## Overview of Project
### Purpose
Green energy stocks were analyzed in order to determine the total volume and yearly return for each stock. A series of code were written to automate and to generate this analysis. Thus, the purpose of this project is to refactor the codes in order to run the analysis in a much efficient manner. By doing so, this will allow Steve to expand the dataset by including the entire stock market over the last years and therefore help him determine green energy stocks that are worth investing in to. 

## Results
### Analysis
For this stock analysis, a series of code needed to be written and since we are working with a big data set, the codes can be quite overwhelming and complicated. Hence, for the first step of this anaylsis, it is imperative to create an outline to organize the codes. After creating the outline, this can be utlized to create codes for each step of the analysis. Please see below for the breakdown of the outline and the codes written for each step.

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
### Results 
For the year 2017, the analysis revealed that almost all of the green energy stocks had a positive rate of return except for the ticker TERP. In addition, the ticker DQ had the greatest return out of all the stocks in the dataset. For the year 2018, however, the roles were reversed. Only two tickers, RUN and ENPH have a positive rate of return and the rest of the stocks did not do well for that year.  
![VBA 2017 Analyis](https://github.com/kntln/stock-analysis/blob/main/VBA_StockAnalysis_2017.png)

![VBA 2017 Analyis](https://github.com/kntln/stock-analysis/blob/main/VBA_StockAnalysis_2018.png)

## Summary
### Advantages and Disadvantages of Refactoring Code
One major advantage of refactoring code is it allows the code to run much efficiently, therefore, more information can be analyzed but with less amount of time. Another advantage of refactoring the code is that it enables the code to be applied generally and is not limited to the specific year of the data set. Lastly, refactoring the code makes the program well-structed and organized. However, refactoring code also comes with limitations. One obvious limitation is it does not allow for customization or speciation for certain functions. In addition, refactoring codes can be sensitive to changes in format of the dataset which can make the codes ecounter error. Lastly, refactoring code are applied to series of standardized basic actions and might be difficult to apply to much complicated functions. 

### Advantages and Disadvantages of the Original and Refactored VBA Script
The advantage of the original VBA script is it allows for customization. For instance, if there is a specific function that Steve wants to accomplish for a specific year, the orginal script can accommodate that. On the other hand, the original script is not efficient in running the analysis. Therefore, the refactored VBA's greatest advantage is its efficiency. The refactored analysis took less than a second to run for both 2018 and 2017, whereas the original script took more than one second to run for both years. Attached below are the macro run times for refactored VBA script.

![VBA 2017 Screenshot](https://github.com/kntln/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA 2018 Screenshot](https://github.com/kntln/stock-analysis/blob/main/VBA_Challenge_2018.png)
