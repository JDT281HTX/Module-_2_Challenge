Sub StockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim i As Long
    Dim startRow As Long
    Dim ticker As String
    Dim startOpen As Double
    Dim endClose As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim currentTicker As String
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    
    For Each ws In ThisWorkbook.Worksheets
        
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
           
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
           
            ws.Cells(1, 8).Value = "Ticker"
            ws.Cells(1, 9).Value = "Quarterly Change"
            ws.Cells(1, 10).Value = "Percent Change"
            ws.Cells(1, 11).Value = "Total Stock Volume"
            
            
            outputRow = 2
            
           
            startRow = 2
            currentTicker = ws.Cells(startRow, 1).Value
            
            
            maxPercentIncrease = -999999
            maxPercentDecrease = 999999
            maxVolume = 0
            
            For i = 2 To lastRow
                
                If ws.Cells(i, 1).Value <> currentTicker Or i = lastRow Then
                   
                    If i = lastRow Then i = i + 1
                    
                    
                    startOpen = ws.Cells(startRow, 3).Value
                    endClose = ws.Cells(i - 1, 6).Value
                    totalVolume = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(i - 1, 7)))
                    quarterlyChange = endClose - startOpen
                    If startOpen <> 0 Then
                        percentChange = (quarterlyChange / startOpen) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    
                    ws.Cells(outputRow, 8).Value = currentTicker
                    ws.Cells(outputRow, 9).Value = Round(quarterlyChange, 2)
                    ws.Cells(outputRow, 10).Value = Round(percentChange, 2)
                    ws.Cells(outputRow, 11).Value = totalVolume
                    
                    
                    If percentChange > maxPercentIncrease Then
                        maxPercentIncrease = percentChange
                        maxPercentIncreaseTicker = currentTicker
                    End If
                    
                    If percentChange < maxPercentDecrease Then
                        maxPercentDecrease = percentChange
                        maxPercentDecreaseTicker = currentTicker
                    End If
                    
                    If totalVolume > maxVolume Then
                        maxVolume = totalVolume
                        maxVolumeTicker = currentTicker
                    End If
                    
                    
                    outputRow = outputRow + 1
                    startRow = i
                    currentTicker = ws.Cells(i, 1).Value
                End If
            Next i
            
            
            
            ws.Cells(1, 14).Value = "Ticker"
            ws.Cells(1, 15).Value = "Value"
            
            ws.Cells(2, 13).Value = "Greatest % Increase"
            ws.Cells(2, 14).Value = maxPercentIncreaseTicker
            ws.Cells(2, 15).Value = Round(maxPercentIncrease, 2)
            
            ws.Cells(3, 13).Value = "Greatest % Decrease"
            ws.Cells(3, 14).Value = maxPercentDecreaseTicker
            ws.Cells(3, 15).Value = Round(maxPercentDecrease, 2)
            
            ws.Cells(4, 13).Value = "Greatest Total Volume"
            ws.Cells(4, 14).Value = maxVolumeTicker
            ws.Cells(4, 15).Value = maxVolume
        End If
    Next ws
End Sub