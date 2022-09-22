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

`Sub AllStocksAnalysis()
    Dim startTime As String
    Dim endTime As String
    
    yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    
    '1) Format the output sheet on the "All Stock Analysis" worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        'create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
    '2) Initialize an array of all tickers
    
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
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    
    '3b) Activate the data worksheet
    Worksheets(yearValue).Activate

    
    
    '3c) Find the numer of rows to loop over
      RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    '4) Loop through the tickers
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
    
    '5) Loop through rows in the data
    Worksheets(yearValue).Activate
        For j = 2 To RowCount
    
    '5a) Find the total volumefor the current ticker
    If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
      End If
    
    
    '5b) Find the stating price for the currnelt ticker
    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        startingPrice = Cells(j, 6).Value
      End If
    
    
    '5c) Find the ending pricefor the current ticker
    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        endingPrice = Cells(j, 6).Value
       End If
    
    Next j
    
    
    '6) Output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
  
`Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year" & (yearValue)
    

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'making cells A3 to C3 bold
    Range("A3:C3").Font.Bold = True
    'making a border A3 to C3
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'making B4:B15 with commas
    Range("B4:B15").NumberFormat = "#,##0"
    'making %
    Range("C4:C15").NumberFormat = "0.0%"
    'autofit column width
    Columns("B").AutoFit
    
        If Cells(4, 3) > 0 Then
         'color cell green
         Cells(4, 3).Interior.Color = vbGreen
        
         ElseIf Cells(4, 3) < 0 Then
            'color cell red
            Cells(4, 3).Interior.Color = vbRed
    
         Else
            'clear cell color
            Cells(4, 3).Interior.Color = xlNone
    
     End If
                    
    dataRowStart = 4
    dataRowEnd = 15
    
        For i = dataRowStart To dataRowEnd
         If Cells(i, 3) > 0 Then
            'change cell color green
            Cells(i, 3).Interior.Color = vbGreen
        
        ElseIf Cells(i, 3) < 0 Then
            'change cell to red
            Cells(i, 3).Interior.Color = vbRed
        
        Else
            'clear cell color
            Cells(i, 3).Interior.Color = xlNone
        
        End If
        
    Next i


`End Sub

When I run the code, the message below shows the runtime

<img width="223" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110318652/191645242-1e30c85c-15ed-48ea-86eb-73c08aff1ccd.png">

### Refactored Code
Like the original set, I formatted the output sheet on the “All Stocks Analysis” worksheet and initialized an array of all tickers, activated the data worksheet to the active year, and got the number of rows to loop over. I start to diverge from the original code by creating my ticker index and setting it to 0. I also created three output arrays for ticker volumes, starting, and ending price. I set the ticker volume as a long while the starting and ending price as doubles. Continuing, created a for loop to initialize the ticker volume to zero, looped over all the rows in the spreadsheet, increased the volume for current ticker, and checked if the current row was the first row with the selected ticker index. I also checked if the current is the last row with the selected ticker, and if the next row’s ticker did not match, I increased the ticker index. I then looped through my arrays to output the ticker, total daily volume, and return. Finally ending the code by formatting the cells and adding a button to put the selected year.
