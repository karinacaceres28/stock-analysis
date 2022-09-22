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

`Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    
    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
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
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
     Dim tickerIndex As Integer
     tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        
        Next i
        
        
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
  
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
          If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1, 1) <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
           
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
        'End If
    
Next i

    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub'



When I run the code, the message below shows the runtime

<img width="261" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/110318652/191646623-e34a664d-f3b5-4d3b-a567-aadbe07ba547.png">

### Original Code Vs Refactored Code
The refactored code is more efficient compared to the original code. The original code has a nested loop, while in the refactored one the code stays in the same loop. Rather than create a nested loop it simply added a separate for loop the results are then obtained in the selected year. The code in the refactored loop is also a lot faster than the original code.

## Summary

### Advantages and Disadvantages of Refactoring Code in General
An advantage of using refactoring code in general is that it is more efficient and allows for data to be pulled on a larger data set. A disadvantage of refactoring code is I see the potential of it being time consuming, this could be due to trying to figure out the purpose of the code and its functionality.

### Refactoring the original VBA script
As shown above, refactoring the original code allowed the code to run a lot faster than it did for the original code. However, the disadvantage was that it was a lot more time consuming due to me trying to figure out the purpose and functionality of the code. The original code was a lot easier for me to complete in the module because it was easier to follow the logic of the code. However, the disadvantage of the original code was it only allows data to be pulled from a smaller set.
