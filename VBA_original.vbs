Sub DQAnalysis()
    Dim startingPrice As Double
    Dim endingPrice As Double

Worksheets("DQ_Analysis").Activate
        
    Range("A1").Value = "DAQO (Ticker: DQ)"
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

Worksheets("2018").Activate
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0

For i = rowStart To rowEnd
    
    If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
        'set starting price
        startingPrice = Cells(i, 6).Value
    End If
    
    If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
        'set ending price
        endingPrice = Cells(i, 6).Value
    End If
    
    'increase totalVolume if ticker is "DQ"
    If Cells(i, 1).Value = "DQ" Then
        totalVolume = totalVolume + Cells(i, 8).Value
    End If

Next i

Worksheets("DQ_Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1

End Sub

Sub AllStocksAnalysis()
    'Find number of rows (before both loops)
    Dim startTime As Single
    Dim endTime As Single

    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer

'1. Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All_Stocks_Analysis").Activate
        
    Range("A1").Value = "All_Stocks_(" + yearValue + ")"
     'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2. Initialize an array of all tickers.
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
    
'3. Prepare for the analysis of tickers.
'   3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double

'   3b) Activate the data worksheet.
    Worksheets(yearValue).Activate

'   3c) Find the number of rows to loop over.
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    

'4. Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0

        '5. Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To rowEnd
            '   5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
                
            '   5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            '   5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j
    
        '6. Output the data for the current ticker.
        Worksheets("All_Stocks_Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        If Cells(4 + i, 3).Value > 0 Then
            'color the cell green
            Cells(4 + i, 3).Interior.Color = vbGreen
        ElseIf Cells(4 + i, 3).Value < 0 Then
            'color the cell red
            Cells(4 + i, 3).Interior.Color = vbRed
        Else
            'clear the cell color
            Cells(4 + i, 3).Interior.Color = x1None
        End If
    Next i
        
'Formatting
    Worksheets("All_Stocks_Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
        
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub ClearWorksheet()
    Cells.Clear
End Sub




