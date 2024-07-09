Attribute VB_Name = "Module1"
Sub Quaterly_change()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim ticker As String
    Dim currentTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim i As Long
    
   For Each ws In ThisWorkbook.Worksheets
        If Application.WorksheetFunction.CountA(ws.Columns(1)) > 1 Then
            ' Assuming the data starts from row 2 and columns are in a specific order
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            outputRow = 2
            
                ticker = ws.Cells(2, 1).Value
                openPrice = ws.Cells(2, 3).Value
                closePrice = ws.Cells(2, 6).Value
                totalVolume = ws.Cells(2, 7).Value
                
            'output headers in currentsheet
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Volume"
            
              ' Initialize the first ticker
            currentTicker = ws.Cells(2, 1).Value
            openPrice = ws.Cells(2, 3).Value
            totalVolume = 0
             For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
            
             ' Check if we have moved to a new ticker
                If ticker <> currentTicker Then
             ' Calculate the changes for the previous ticker
                    closePrice = ws.Cells(i - 1, 6).Value
                    quarterlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                      percentChange = ((closePrice - openPrice) / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
              ' Output the previous ticker's data
                    ws.Cells(outputRow, 9).Value = currentTicker
                    ws.Cells(outputRow, 10).Value = quarterlyChange
                    ws.Cells(outputRow, 11).Value = percentChange
                    ws.Cells(outputRow, 12).Value = totalVolume
                    outputRow = outputRow + 1
                    
                 ' Initialize new ticker data
                    currentTicker = ticker
                    openPrice = ws.Cells(i, 3).Value
                    totalVolume = 0
                End If
                 ' Update the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Next i
            
            ' Output the last ticker's data
            closePrice = ws.Cells(lastRow, 6).Value
            quarterlyChange = closePrice - openPrice
            If openPrice <> 0 Then
             percentChange = ((closePrice - openPrice) / openPrice) * 100
     Else
                percentChange = 0
            
                End If
            ws.Cells(outputRow, 9).Value = currentTicker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = totalVolume
            
      End If
    Next ws
End Sub

