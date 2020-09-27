# An Analysis on Green Stocks Performance

## Overview of Project

### The analysis aims to show the performance of Green Energy stocks in the stock market over a period of time.  The results of the analysis will help the client, Steve, advise his parents if they should stay invested in DAQO New Energy Corp, and advise them on how to diversify their portfolio.

## Results

### A macro called ***AllStocksAnalysis*** was created to calculate the total volumne and yearly return of any given stock in the dataset, for the desired year.

### The following code allowed Steve to enter the year he wants to run the analysis on:
```
yearValue = InputBox("What year would you like to run the analysis on?")
```
### The yearValue result was also used in the macro to determine which worksheet will be used for the code.  A nested loop was created that told Excel to loop through the whole spreadsheet, looking at all tickers & adding the total volume, as long as it was the same ticker.  Within this loop, the code determined the starting price and closing price of the current ticker.  

### Here is part of that code:
```
'4) Loop through tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
            
            '5) loop through rows in the data
           Worksheets(yearValue).Activate
           For j = 2 To RowCount
           
           '5a) Get total volume for current ticker
               If Cells(j, 1).Value = ticker Then
    
                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value
    
               End If
               
               '5b) get starting price for current ticker
               If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
                startingPrice = Cells(j, 6).Value
    
               End If
               
               '5c) get ending price for current ticker
               If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
                endingPrice = Cells(j, 6).Value
    
               End If
    
           Next j
```

### Below are the results for 2017 and 2018.

[2017 All Stocks Analysis](https://github.com/ABonuan/stock-analysis/blob/master/2017%20All%20Stocks%20Analysis.png)

[2018 All Stocks Analysis](https://github.com/ABonuan/stock-analysis/blob/master/2018%20All%20Stocks%20Analysis.png)

### A refactored version of the macro was also created to optimize the code.  This version removed the nested loops which went through the whole spreadsheet before adding and assigning the the total volume to a variable.  A separate `for` loop was created to initialize each ticker's total volume first.  The refactored code then, made use of arrays, which stored data from each row, and a specific index to access those arrays, which made it more efficient. 

### Here is a part of the refactored code:
```
 '1a) Create a ticker Index
    tickerIndex = 0
       
    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    Worksheets(yearValue).Activate
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        
    Next i
        
          
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1
                
            End If
                       
    Next i
```

### Below are the results for 2017 and 2018, using the refactored code.

[VBA Challenge 2017](https://github.com/ABonuan/stock-analysis/blob/master/VBA%20Challenge%202017.png)

[VBA Challenge 2018](https://github.com/ABonuan/stock-analysis/blob/master/VBA%20Challenge%202018.png)

### Each analysis ran 8 to 9 times faster on the refactored code.

### Based on the dataset and the analysis on the stocks themselves, it seems like Green Energy stocks are unpredictable.  Take DAQO for example.  In 2017, total volume traded was 35,796,200 and the yearly return was 199.4%.  In 2018, DAQO's total volume was 107,873,900 and yearly return was -62.6%.  It was traded more than the previous year, but lost value.  Based on this dataset, the stocks that were more consistently traded and remained in the positive in yearly return were ENPH & RUN.


## Summary

### Refactoring code can lead to more efficient & readable code, especially if the dataset increases in volume.  Refactoring can become tricky though, when the original code is very complicated and not as readable, therefore it is difficult to make sense of the logic.

### In this macro, it made sense to refactor the code.  Steve wants to be able to analyze all stocks over the past few years.  Just with the given dataset, the analysis improved in speed.  As the dataset grows, I would imagine the code running slower though.  We may need to use another tool altogether for datasets with really high volume.

