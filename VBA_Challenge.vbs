Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    'Ask the user to to input a year they would like to review
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Begin timer after the input data is received for better representation of time it takes for the subroutine to run
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
    
    'Create a ticker Index
    tickerIndex = 0

    'Create three output arrays to hold values for volumes, starting prices, and ending prices
    Dim tickerVolumes(12) as Long
    Dim tickerStartingPrices(12) as Single
    Dim tickerEndingPrices(12) as Single
    
    'Create a for loop to initialize the tickerVolumes to zero. 
    For i = 1 To 12 
        tickerVolumes(tickerIndex) = 0

        'Activate the worksheet with the pertinent data and loop over all the rows. 
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            
            'Find cells with the current ticker and increase tickerVolume.
            If Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
            End If
            
            'Check if the current row is the first row with the selected ticker and store in tickerStartingPrices array.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
            
            'Check if the current row is the last row with the selected ticker and store in tickerEndingPrices array.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                
                'Increase the tickerIndex to cycle through other tickers.
                tickerIndex = tickerIndex + 1

            End If
        
        Next j
        
    Next i
    
    'Loop through your arrays and pull data to fill worksheet with data for Ticker, Total Daily Volume, and Return.
    For k = 0 To 11
        
        'Activate the worksheet for displaying the data
        Worksheets("All Stocks Analysis").Activate

        Cells(4 + k, 1).Value = tickers(k)
        Cells(4 + k, 2).Value = tickerVolumes(k)
        Cells(4 + k, 3).Value = (tickerEndingPrices(k) / tickerStartingPrices(k)) - 1
        
    Next k
    
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

    'Display time that it took for subroutine to run
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub