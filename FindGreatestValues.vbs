Attribute VB_Name = "Module2"
Sub FindGreatestValues()

    Dim ws As Worksheet
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    
    greatestPercentIncrease = 0
    greatestPercentDecrease = 0
    greatestTotalVolume = 0
    
    For Each ws In Worksheets
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            ' Update variables for tracking greatest values
            If ws.Cells(i, 11).Value > greatestPercentIncrease Then
                greatestPercentIncrease = ws.Cells(i, 11).Value
                greatestPercentIncreaseTicker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 11).Value < greatestPercentDecrease Then
                greatestPercentDecrease = ws.Cells(i, 11).Value
                greatestPercentDecreaseTicker = ws.Cells(i, 9).Value
            End If
            
            If ws.Cells(i, 12).Value > greatestTotalVolume Then
                greatestTotalVolume = ws.Cells(i, 12).Value
                greatestTotalVolumeTicker = ws.Cells(i, 9).Value
            End If
        
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 16).Value = greatestPercentIncreaseTicker
            ws.Cells(2, 17).Value = greatestPercentIncrease
            ws.Cells(3, 16).Value = greatestPercentDecreaseTicker
            ws.Cells(3, 17).Value = greatestPercentDecrease
            ws.Cells(4, 16).Value = greatestTotalVolumeTicker
            ws.Cells(4, 17).Value = greatestTotalVolume
        Next i
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    Next ws

End Sub
