Sub summarizeStock()

    Dim strTicker As String
    Dim lastTicker As String
    Dim intTickerCount As Integer
    Dim dblOpen As Double
    Dim dblClose As Double
    Dim dblChange As Double
    Dim dblPercentChange As Double
    Dim lngVolume As LongLong
    
    'variables for navigating spreadhsheet
    Dim ws As Worksheet
    Dim lngLastRow As Long
    Dim intLastCol As Integer
    Dim r
            
    'loop through each worksheet
    For Each ws In Worksheets
    
        lngLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        intLastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
                
        'initialize
        lastTicker = ws.Cells(2, 1).Value
        intTickerCount = 1
        lngVolume = 0
        dblOpen = 0
        dblChange = 0
        dblPercentChange = 0
                
        'set the table headers
        ws.Cells(intTickerCount, intLastCol + 2).Value = "Ticker"
        ws.Cells(intTickerCount, intLastCol + 2).Font.Bold = True
        ws.Cells(intTickerCount, intLastCol + 3).Value = "Yearly Change"
        ws.Cells(intTickerCount, intLastCol + 3).Font.Bold = True
        ws.Cells(intTickerCount, intLastCol + 4).Value = "Percent Change"
        ws.Cells(intTickerCount, intLastCol + 4).Font.Bold = True
        ws.Cells(intTickerCount, intLastCol + 5).Value = "Total Stock Volume"
        ws.Cells(intTickerCount, intLastCol + 5).Font.Bold = True
        
        
        'navigate the rows, 2 through row count
        For r = 2 To lngLastRow
        
            strTicker = ws.Cells(r, 1).Value
            
            If strTicker <> lastTicker Then
                
                'grab the closing price (last value)
                'if the stock never opened, thanks PLNT/2014, put zeros...
                If dblOpen = 0 Then
                    dblClose = 0
                    dblChange = 0
                    dblPercentChange = 0
                Else
                    dblClose = ws.Cells(r - 1, 6).Value
                    dblChange = dblClose - dblOpen
                    dblPercentChange = dblChange / dblOpen
                End If
                
                'set table values
                ws.Cells(intTickerCount + 1, intLastCol + 2).Value = lastTicker
                ws.Cells(intTickerCount + 1, intLastCol + 3).Value = dblChange
                ws.Cells(intTickerCount + 1, intLastCol + 4).Value = dblPercentChange
                ws.Cells(intTickerCount + 1, intLastCol + 5).Value = lngVolume
                            
                'format cells
                If dblChange >= 0 Then
                    ws.Cells(intTickerCount + 1, intLastCol + 3).Interior.ColorIndex = 4
                Else
                    ws.Cells(intTickerCount + 1, intLastCol + 3).Interior.ColorIndex = 3
                End If
                ws.Cells(intTickerCount + 1, intLastCol + 4).NumberFormat = "0.00%"
                
                'increment ticker counter
                intTickerCount = intTickerCount + 1
                                
                'reset variables
                lngVolume = 0
                dblChange = 0
                dblPercentChange = 0
                dblOpen = 0
                lastTicker = strTicker
                   
            End If
                        
            'set the open at the first value for the stock
            If dblOpen = 0 And ws.Cells(r, 3).Value <> 0 Then
                dblOpen = ws.Cells(r, 3).Value
            End If
            
            'increment the stock volume
            lngVolume = lngVolume + ws.Cells(r, 7).Value
    
        Next r
        

        'get the last row
        ws.Cells(intTickerCount + 1, intLastCol + 2).Value = strTicker
        ws.Cells(intTickerCount + 1, intLastCol + 3).Value = dblChange
        ws.Cells(intTickerCount + 1, intLastCol + 4).Value = dblPercentChange
        ws.Cells(intTickerCount + 1, intLastCol + 5).Value = lngVolume
        
        'format cells
        If dblChange >= 0 Then
            ws.Cells(intTickerCount + 1, intLastCol + 3).Interior.ColorIndex = 4
        Else
            ws.Cells(intTickerCount + 1, intLastCol + 3).Interior.ColorIndex = 3
        End If
        ws.Cells(intTickerCount + 1, intLastCol + 4).NumberFormat = "0.00%"
        
        'autosize
        ws.Range("I1:L1").Columns.AutoFit
        
    Next ws

End Sub