# Stock Analysis using Excel VBA


## Overview of Project
We performed an analysis on the green stock dataset to determine the price movement of the stocks, to help investors chose the appropriate stocks for trading. Investors are interested in DAQO energy stock and wanted to evaluate the current trends of that stock.  We used Visual Basic for applications (VBA) in Excel to help perform an automated analysis for DAQO stock as well as thousands of other stocks for future use. 


### Purpose

In this analysis, we examine the trends in several green energy stocks from two years to evaluate the investment in DAQO New energy stock. 


## Results

A new worksheet was created in excel that preresented the result of the analysis of the green stock dataset. The results include Tickers, Total Volume and the yearly returns. At first, we evaluated the DAQO stocks trade and yearly return in 2018, followed by comparing it with the rest of the stocks to find alternative choices. The results showed that the price of DAQO stocks had dropped by 63% by the end of the year 2018. The results of the analysis for all the stocks in the year 2018 are shown below that the investor can get with a click of a button.


![Picture2](https://user-images.githubusercontent.com/79213116/116793082-99f74b80-aa92-11eb-9f3d-24c2d1d875a1.png)




Following these results, we refactored the code to get the results for the entire stock market for last few years. Code refactoring helps provide the results of the analysis faster. The refarctored code only took 0.093 seconds to run the analysis for the year 2017 and 0.085 for the year 2018. 
The refactored code in the VBA is presented below along with the pictures of the elapsed runtime after refoacotring.

code:

```

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Sart counting the time
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
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
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
    
    'stop the timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```


![Picture2](https://user-images.githubusercontent.com/79213116/116793780-dc228c00-aa96-11eb-8b12-21bc60d8f113.png)


![Picture2](https://user-images.githubusercontent.com/79213116/116793814-10964800-aa97-11eb-8eac-aee33fc57f32.png)


### Summary 
Adavantages of refractored coding:

- it makes the code more readable
- it reduces the number of lines of code
- it improves the speed and performance of the program

Disadvantages of refractored coding:

- it can be very time consuming, in case of a mistake it may take a lot of time to resolve due to the complexity of the code.

After the use of refarctored code on the green stock analysis, the run time was reduced from approximately 32 seconds to the time shown in the pictures above. 
Link for the detailed analysis can be found here: [VBA_CHALLENGE](https://github.com/Komal77rao/stock-analysis/blob/main/VBA_challenge.xlsm)
