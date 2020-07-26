Attribute VB_Name = "Module2"
Sub AllStocksAnalysisRefactored()

    yearValue = InputBox("What year would you like to run the analysis on?")
    startTimer = Timer
    
    '1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    '2) Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '3) Initialize array of all tickers
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
    
    '4a) Activate data worksheet
    Worksheets(yearValue).Activate
    
    '4b) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '5a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0
    

    '5b) Create three output arraysDim tickerVolume As Long
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingprices(12) As Single
    
    
    
    '6a) Initialize ticker volumes to zero
   
   For i = 0 To 11
    tickerVolume(i) = 0
   
   Next i
        
       
   
    
        
    '6b) loop over all the rows
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
        totalVolume = tickerVolume(tickerIndex)
         tickerVolume(tickerIndex) = totalVolume + Cells(i, 8).Value
         
                
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 7).Value
            
        End If
            
            
        'End If
        
        '7c) check if the current row is the last row with the selected ticker
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingprices(tickerIndex) = Cells(i, 7).Value
            
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
            
            
            
            
        'End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    '9) Formatting
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

End Sub
