Attribute VB_Name = "Module1"
Sub CalculateStockSummary()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryTableRow As Integer
    Dim ticker As String
    Dim yearlyChange As Double
    Dim firstOpen As Double
    Dim lastClose As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    For Each ws In Worksheets
        summaryTableRow = 2
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        firstOpen = ws.Cells(2, 3).Value ' Opening price for the first row
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Calculate yearly change, percent change, and total volume
                ticker = ws.Cells(i, 1).Value
                lastClose = ws.Cells(i, 6).Value
                yearlyChange = lastClose - firstOpen
                If firstOpen <> 0 Then
                    percentChange = yearlyChange / firstOpen
                Else
                    percentChange = 0
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Output the results to the summary table
                ws.Cells(summaryTableRow, 9).Value = ticker
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 12).Value = totalVolume
                
                ' Move to the next row in the summary table
                summaryTableRow = summaryTableRow + 1
                
                ' Reset variables for the next ticker
                yearlyChange = 0
                firstOpen = ws.Cells(i + 1, 3).Value ' Update opening price for the next ticker
                totalVolume = 0
            Else
                ' Accumulate total volume for the same ticker
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        For i = 2 To lastRow
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            ws.Cells(i, 11).NumberFormat = "0.00%"
        Next i
    Next ws

End Sub

