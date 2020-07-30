Sub test():
    
    Dim i As Long
    Dim total As Double
    Dim lastRowA As Long
    
    For Each ws In Worksheets
    
        lastRowA = ws.Cells(Rows.Count, "A").End(xlUp).Row
        total = 0
        
        'adds headers to columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'x will represent the parent row of each ticker value
        x = 2
        'y will represent the row in "Ticker, yearly change, percent change, total" for each ticker
        y = 2
    
        'for loop goes through every row in column A
        For i = 3 To lastRowA
            'if the ticker in row i is equal to the ticker in the previous row
            If (ws.Cells(i, 1).Value = ws.Cells((i - 1), 1).Value) Then
                'adds the newest volume number to the running total for that ticker
                total = total + ws.Cells(i, 7).Value
            'if the ticker value in the current row is not equal to the previous ticker (meaning we're now on a new ticker)
            Else:
                'moves the value of x to the row of the new ticker value
                x = i
                'moves the placement of the values for "Ticker, yearly change, percent change, total" one row down
                y = y + 1
                'restarts the total back to 0 because we're counting a new ticker now
                total = 0
                'adds the newest volume number to the running total for that ticker
                total = total + ws.Cells(i, 7).Value
            End If
            'the ticker value in I2 will be/continue to be the ticker value of A,x
            ws.Cells(y, 9).Value = ws.Cells(x, 1).Value
            'the yearly change value in J,y will be the current row's closing price - the first ticker of the group's opening price
            ws.Cells(y, 10).Value = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value
            'avoids problems with dividing by 0
            If ws.Cells(x, 3).Value <> 0 Then
                'the percent change value in K,y will be the (yearly change value / the first ticker of the group's opening price) * 100
                ws.Cells(y, 11).Value = Round(((ws.Cells(y, 10).Value / ws.Cells(x, 3).Value) * 100), 2)
            End If
            'the total stock volume value in L,y will be the running total of all the volume for that ticker
            ws.Cells(y, 12).Value = total
        Next i
        
        'changes color of yearly change cells depending if they're positive or negative
        For i = 2 To lastRowA
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        Dim greatestPercent As Double
        Dim leastPercent As Double
        Dim greatestVolume As Double
        Dim lastRowK As Integer
        
        lastRowK = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        'adds headers to new table
        ws.Range("n2").Value = "Greatest % Increase"
        ws.Range("n3").Value = "Greatest % Decrease"
        ws.Range("n4").Value = "Greatest Total Volume"
        ws.Range("o1").Value = "Ticker"
        ws.Range("p1").Value = "Value"
        
        'starting values to compare other values to
        greatestPercent = 0
        leastPercent = 0
        greatestVolume = 0
        
        For i = 2 To lastRowK
            'if the value of yearly change at row i is greater than the current value of greatestPercent
            If ws.Cells(i, 11).Value > greatestPercent Then
                'update the value of greatestPercent to the new highest value
                greatestPercent = ws.Cells(i, 11).Value
                'update the ticker in the table to the ticker of the new highest value
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                'update the value in the table to the new highest value
                ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
            End If
            'same conditional as before, but with the lowest value
            If ws.Cells(i, 11).Value < leastPercent Then
                leastPercent = ws.Cells(i, 11).Value
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
            End If
            'same conditional as before, but with the greatest total volume
            If ws.Cells(i, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
            End If
        Next i
        
    Next ws
    
End Sub