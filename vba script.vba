Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim rowStart As Integer
    Dim tickerRow As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String

    For Each ws In ThisWorkbook.Worksheets
        tickerRow = 2
        totalVolume = 0
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        maxIncreaseTicker = ""
        maxDecreaseTicker = ""
        maxVolumeTicker = ""
        rowStart = 2
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Add titles to the columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Initialize start price
        startPrice = ws.Cells(2, 3).Value

        ' Loop through all rows
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                endPrice = ws.Cells(i, 6).Value

                yearlyChange = endPrice - startPrice
                If startPrice <> 0 Then
                    percentChange = (yearlyChange / startPrice) * 100
                Else
                    percentChange = 0
                End If

                ' Check for max and min increase/decrease and max volume
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ticker
                End If
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ticker
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    maxVolumeTicker = ticker
                End If

                ' Output the data
                ws.Cells(tickerRow, 9).Value = ticker
                ws.Cells(tickerRow, 10).Value = yearlyChange
                ws.Cells(tickerRow, 11).Value = percentChange
                ws.Cells(tickerRow, 11).NumberFormat = "0.00%"
                ws.Cells(tickerRow, 12).Value = totalVolume

                ' Conditional formatting
                If yearlyChange > 0 Then
                    ws.Cells(tickerRow, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(tickerRow, 10).Interior.Color = vbRed
                End If

                ' Reset the total volume and start price
                totalVolume = 0
                tickerRow = tickerRow + 1
                If i + 1 <= lastRow Then
                    startPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Output the greatest % increase, % decrease, and total volume
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(2, 17).Value = maxIncreaseTicker
        ws.Cells(3, 17).Value = maxDecreaseTicker
        ws.Cells(4, 17).Value = maxVolumeTicker
        ws.Cells(2, 18).Value = maxIncrease
        ws.Cells(3, 18).Value = maxDecrease
        ws.Cells(4, 18).Value = maxVolume
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        ws.Cells(4, 18).NumberFormat = "#,##0"

    Next ws

    MsgBox "Stock Market Analysis Complete!"
End Sub

