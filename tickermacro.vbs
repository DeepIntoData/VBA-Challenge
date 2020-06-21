Sub tickermacro()

'declare variables
Dim Op As Double 'Open Price
Dim Ep As Double 'End Price
Dim Vol As Double 'Volume
Dim t As Integer 'Counter Value
Dim YC As Double 'Yearly Change Variable

'Sheet Loop
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate

'Set titles
Cells(1, 9).Value = "TickerName"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

'Set First Row (Position Reference)
    t = 2
    Cells(t, 9).Value = Cells(t, 1).Value
    Op = Cells(2, 3).Value

'Determine Last Row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox lastrow 'test

Columns("K").NumberFormat = "0.00%"

'Capture the row number of the next ticker (THIS I SPECIFIC TO TICKERNAME, -1 for dates)
        For i = 2 To lastrow
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                'Volume
                    Vol = Vol + Cells(i, 7).Value
                    Cells(t, 12).Value = Vol
                
                'End Price
                    Ep = Cells(i, 6).Value
                
                'Yealry Change
                    YC = Ep - Op
                    Cells(t, 10) = YC
                
                'Color 
                If YC > 0 Then
                    Cells(t, 10).Interior.Color = vbGreen
                ElseIf YC < 0 Then
                        Cells(t, 10).Interior.Color = vbRed
                End If
                
                'Percent Change
                If Op = 0 Then
                    Cells(t, 11).Value = 0
                Else
                    Cells(t, 11).Value = ((Ep - Op) / Op)
                End If
                
                t = t + 1
              
                'Ticker
                Cells(t, 9).Value = Cells(i + 1, 1).Value
                
                'Reset Volume counter
                Vol = 0
                Else
                'Inverse
                Vol = Vol + Cells(i, 7).Value
                         
    End If
    
Next i

Next ws

End Sub
