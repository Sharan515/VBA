Sub tickerStock()

    ' Loop through each worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
        ' Find the last row of the table
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

        ' Add headers in columns I to L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables
        Dim openPrice As Double
        Dim closePrice As Double
        Dim quarterlyChange As Double
        Dim ticker As String
        Dim percentChange As Double
        Dim volume As Double
        Dim row As Long
        Dim column As Long
        
        ' Initialize variables
        volume = 0
        row = 2
        column = 1
        
        ' Set the initial open price
        openPrice = ws.Cells(2, column + 2).Value
        
        ' Loop through stock data
        For i = 2 To lastRow
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ' Set ticker name
                ticker = ws.Cells(i, column).Value
                ws.Cells(row, column + 8).Value = ticker
                
                ' Set close price
                closePrice = ws.Cells(i, column + 5).Value
                
                ' Calculate quarterly change
                quarterlyChange = closePrice - openPrice
                ws.Cells(row, column + 9).Value = quarterlyChange
                
                ' Calculate percent change
                percentChange = quarterlyChange / openPrice
                ws.Cells(row, column + 10).Value = percentChange
                ws.Cells(row, column + 10).NumberFormat = "0.00%"
                
                ' Calculate total volume per quarter
                volume = volume + ws.Cells(i, column + 6).Value
                ws.Cells(row, column + 11).Value = volume
                
                ' Move to next row
                row = row + 1
                
                ' Reset open price to next ticker
                openPrice = ws.Cells(i + 1, column + 2).Value
                
                ' Reset volume for next ticker
                volume = 0
            Else
                volume = volume + ws.Cells(i, column + 6).Value
            End If
        Next i
        
        ' Find the last row for the ticker column
        Dim quarterlyChangeLastRow As Long
        quarterlyChangeLastRow = ws.Cells(ws.Rows.Count, 9).End(xlUp).row
        
        ' Apply cell colors for quarterly changes
        For j = 2 To quarterlyChangeLastRow
            If ws.Cells(j, 10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10 ' Green
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3 ' Red
            End If
        Next j
        
        ' Set headers for Greatest % Increase, % Decrease, and Total Volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Find and display the greatest percent change, greatest decrease, and highest volume
        For k = 2 To quarterlyChangeLastRow
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quarterlyChangeLastRow)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quarterlyChangeLastRow)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quarterlyChangeLastRow)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            End If
        Next k
        
        ' Make header rows bold and auto-fit columns
        ws.Range("I:Q").Font.Bold = True
        ws.Range("I:Q").EntireColumn.AutoFit
    Next ws

End Sub
