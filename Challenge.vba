Sub tickertime()
'define variables
Dim ticker2 As Long
ticker2 = 2
LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim total_volume As Double
    total_volume = 0
Dim start As Long
    start = 2
Dim percent As Double

'set values

For i = 2 To LR
total_volume = total_volume + Cells(i, 7).Value

'print values
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Cells(ticker2, 9).Value = ticker
        Cells(ticker2, 10).Value = total_volume
       
        'reseting to zero
        total_volume = 0
        
        ychange = Cells(i, 6) - Cells(start, 3).Value
    If Cells(start, 3).Value > 0 Then
        percent = 100 * ychange / Cells(start, 3).Value
    Else
        percent = 0
    End If
    
    If ychange > 0 Then
        Cells(ticker2, 11).Interior.ColorIndex = 4
    ElseIf ychange < 0 Then
        Cells(ticker2, 11).Interior.ColorIndex = 3
    Else
        Cells(ticker2, 11).Interior.ColorIndex = 0
    End If
        
        start = i + 1
        Cells(ticker2, 11).Value = ychange
        Cells(ticker2, 12).Value = percent
        
         ticker2 = ticker2 + 1
    End If
Next i

End Sub
