Sub summaries():
	'Challenge 2 (VBA) by Paul Schnase
    Dim i As Long 'index of row we are checking
    Dim j As Long 'index of row we are writing to
    Dim ticker As String 'ticker of current stock
    Dim vol As LongLong 'running total of volume
    Dim start As Double 'opening price of stock at beginning of year
    Dim bottom As Long 'where the table ends
    Dim best As Double
    Dim worst As Double
    Dim maxVol As LongLong
    
    
    
    'label for showing progress as code runs
    Cells(1, 8).Value = "Current ticker"
    
    'loop over the worksheets
    For Each ws In Worksheets
        'set up column titles for main assignment
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'set up titles for bonus objective
        ws.Cells(2, 13).Value = "Greatest % increase"
        ws.Cells(3, 13).Value = "Greatest % decrease"
        ws.Cells(4, 13).Value = "Greatest total volume"
        ws.Cells(1, 14).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"
        
        'starting values
        ticker = ws.Cells(2, 1).Value
        start = ws.Cells(2, 3).Value
        vol = ws.Cells(2, 7).Value
        ws.Cells(2, 15).Value = 0
        ws.Cells(3, 15).Value = 0
        ws.Cells(4, 15).Value = 0
        
        'find bottom of table
        bottom = ws.Range("A1").End(xlDown).Row
        
        'set row to write to
        j = 2
        'loop over all remaining rows of sheet
        For i = 3 To (bottom + 1)
            'if ticker is the same as previous row, add volume to vol variable
            If ticker = ws.Cells(i, 1).Value Then
                vol = vol + ws.Cells(i, 7).Value
            
            Else
                'otherwise write the ticker symbol, change for the year, percent change and volume
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = (ws.Cells(i - 1, 6).Value - start)

                ws.Cells(j, 11).Value = (ws.Cells(i - 1, 6).Value - start) / start
                ws.Cells(j, 12).Value = vol
                'color price change based on the sign of the change
                If ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                End If
                
                
                If ws.Cells(j, 11).Value > ws.Cells(2, 15) Then
                    'if this stock did better than all previous, record it
                    ws.Cells(2, 14).Value = ticker
                    ws.Cells(2, 15).Value = FormatPercent(ws.Cells(j, 11), 2)
                ElseIf ws.Cells(j, 11).Value < ws.Cells(3, 15).Value Then
                    'if this stock did worse, record that
                    ws.Cells(3, 14).Value = ticker
                    ws.Cells(3, 15).Value = FormatPercent(ws.Cells(j, 11), 2)
                End If
                
                If vol > ws.Cells(4, 15).Value Then
                    'if this stock had higher volume, record that
                    ws.Cells(4, 14).Value = ticker
                    ws.Cells(4, 15).Value = vol
                End If
                
                
                'make the percent change formatted as a percent
                ws.Cells(j, 11).Value = FormatPercent(ws.Cells(j, 11), 2)
                'write to next line
                j = j + 1
                
                'get the next ticker symbol, starting price and first volume
                ticker = ws.Cells(i, 1).Value
                start = ws.Cells(i, 3).Value
                vol = ws.Cells(i, 7).Value
                'display progress on the sheet the macro is run from
                Cells(2, 8).Value = ticker
            End If
            
        Next i
    Next ws
    'when the program is done, clear out the progress tracker
    Cells(1, 8).Clear
    Cells(2, 8).Clear
End Sub

