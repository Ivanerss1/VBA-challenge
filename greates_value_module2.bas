Attribute VB_Name = "Module2"
Sub Greatest_Value()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim totalVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim tickerVolume As Object
    Dim currentTicker As String
    
    
For Each ws In ThisWorkbook.Worksheets
lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    Set tickerVolume = CreateObject("Scripting.Dictionary")
    
    'variables
maxIncrease = -1000000
maxDecrease = 1000000
maxVolume = 0

'loop through each row
For i = 2 To lastRow
    percentChange = ws.Cells(i, 11).Value
    
    ' maximum increae
    If percentChange > maxIncrease Then
    maxIncrease = percentChange
    maxIncreaseTicker = ws.Cells(i, 9).Value
End If
    
        'maximum Decrease
    If percentChange < maxDecrease Then
    maxDecrease = percentChange
    maxDecreaseTicker = ws.Cells(i, 9).Value
End If

        'total volume
    currentTicker = ws.Cells(i, 9).Value
    If tickerVolume.exists(currentTicker) Then
    tickerVolume(currentTicker) = tickerVolume(currentTicker) + ws.Cells(i, 13).Value
    Else
    tickerVolume.Add currentTicker, ws.Cells(i, 13).Value
    End If
Next i

        'Find ticker with max vol.
    For Each currentTicker In tickerVolume.keys
    If tickerVolume(currentTicker) > maxVolume Then
        maxVolume = tickerVolume(currentTicker)
        maxVolumeTicker = currentTicker
End If
        Next currentTicker
        
        'results output
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatesr % Increase"
    ws.Cells(2, 15).Value = maxIncreaseTicker
    ws.Cells(2, 16).Value = Format(maxIncrease, "0.00") & "%"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = maxDecreaseTicker
    ws.Cells(3, 16).Value = Format(maxDecrease, "0.00") & "%"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = maxVolumeTicker
    ws.Cells(4, 16).Value = maxVolume
    
        Next ws
    
        
    


End Sub
