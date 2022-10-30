# Stock Analysis

## Overview of Project
In the Stock Analysis project, we are helping Steve with the analysis of a green stock for his parents to see if it's worth making an investment in. They decided to invest in DAQO New Energy Corp (DQ) so this was the focus of the analysis.
### DQ Analysis
We utilized VBA code to assist Steve in this analysis. We created the below code: 

Sub DQAnalysis()

Worksheets("DQ Analysis").Activate

Range("A1").Value = "DAQO (Ticker:DQ)"

'Create a header row

Cells(3, 1).Value = "Year"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

Worksheets("2018").Activate

'set initial volume to 0
totalVolume = 0

Dim startingPrice As Double
Dim endingPrice As Double

'Establish the number of rows to loop over

rowStart = 2
'DELETE: rowEnd= 3013
'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

'loop over all the rows

For i = rowStart To rowEnd
    'increase totalVolume
    
    If Cells(i, 1).Value = "DQ" Then
    
    'increase totalVolume by the value in the current row
    
    totalVolume = totalVolume + Cells(i, 8).Value
    
    End If
    
    
    If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
    'set starting price
    
    startingPrice = Cells(i, 6).Value
    
    End If
    
    If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
    
    endingPrince = Cells(i, 6).Value
    
    End If


Next i

'MsgBox (totalVolume)

Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    
    
End Sub

This code provided us with the total traded DQ shares in 2018, showing the returns in 2018 were not good. 

### All Stocks Analysis
We then ran an analysis on All Stocks, using the following code to find which stocks had the best return: 

Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer

'1) Format the output sheet on the "All Stocks Analysis" Worksheet
Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'2) Initialize an array of ticklers.
Dim tickers(12) As String

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
'3a) Initialize variables for the starting price and ending price

Dim startingPrice As Single
Dim endingPrice As Single

'3b) Activate the data worksheet

Worksheets("2018").Activate

'3c) Find the number of rows to loop over

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers

For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
'5) Loop through the rows in the data

Worksheets("2018").Activate
For j = 2 To RowCount

'5a) Find total volume for the current ticker
If Cells(j, 1).Value = ticker Then
    totalVolume = totalVolume + Cells(j, 8).Value
    End If
    
'5b) Find starting price for the current ticker
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    startingPrice = Cells(j, 6).Value
    End If
    
'5c) Find ending price for the current ticker
If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    endingPrice = Cells(j, 6).Value
    End If
    
Next j

'6) Output the data for the current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

Next i

    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    

    
End Sub

This provided us with which stocks were a positive return.  


## Results

The first DQ Analysis provided us with the following, showing the total traded DQ shares in 2018 and the return: 
image.png

The All Stocks Analysis provided us with the following, showing the total traded shares per stock in 2018 and the overall return for each. 

image.png

## Summary 

1. What are the advantages or disadvantages of refactoring code?
The advantage of refactoring code is the decrease in screen running time than the original script. Another advantage is it's easier to read & understand. 
A disadvantage is the time it takes to refactor the code and ensure every step is accurate and there are no bugs when the new code runs. 

2. How do these pros and comns apply to refactoring the original VBA script? 
The refactored code was more complex, but the run time was decreased so it was mroe efficient. 

