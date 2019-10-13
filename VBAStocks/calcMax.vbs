Sub calcMax()

    'challenge variables
    Dim largestPercentIncrease As Double
    Dim largestPercentDecrease As Double
    Dim largestVolume As LongLong
    Dim lastRow As Long
    Dim r
    Dim ws

    For Each ws In Worksheets

        lngLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        'O = 15
        'P = 16
        'Q = 17
        
        'setup labels in 8 rows past the last data column, rows 2,3 and 4
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 15).Font.Bold = True
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Font.Bold = True
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Font.Bold = True
        
        'headers for ticker/value, row 1, columns 9 & 10 past the last data col
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 16).Font.Bold = True
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(1, 17).Font.Bold = True
        
        'format %s
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
        'initialize the variables
        largestPercentIncrease = 0
        largestPercentDecrease = 0
        largestVolume = 0
        
        'loop through the summary table and find the largest percent increase/decreas and volume
        For r = 2 To lngLastRow
            
            'compare values to the variables
            'largest percent increase
            If ws.Cells(r, 11).Value > largestPercentIncrease Then
                largestPercentIncrease = ws.Cells(r, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(r, 9).Value
                ws.Cells(2, 17).Value = largestPercentIncrease
            End If
            'largest percent decrease
            If ws.Cells(r, 11).Value < largestPercentDecrease Then
                largestPercentDecrease = ws.Cells(r, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(r, 9).Value
                ws.Cells(3, 17).Value = largestPercentDecrease
            End If
            'biggest volume
            If ws.Cells(r, 12) > largestVolume Then
                largestVolume = ws.Cells(r, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(r, 9).Value
                ws.Cells(4, 17).Value = largestVolume
            End If
        
        Next r
        
        ws.Range("O1:Q1").Columns.AutoFit

    Next ws

End Sub