Sub reset()

    Dim ws As Worksheet
    
    ' Loop through all worksheets in the active workbook
    For Each ws In ActiveWorkbook.Worksheets
        ' Delete columns I to ZZ on the current worksheet
        ws.Range("I:ZZ").Delete
    Next ws
    
    ' Finally, select the "Q1" worksheet
    Worksheets("Q1").Select

End Sub
