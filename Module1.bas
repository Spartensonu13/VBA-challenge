Attribute VB_Name = "Module1"
Sub TickerInfo():

   Dim tickerSymbol As String
   Dim volumeOfStock As Long
   Dim openPrice As Double
   Dim closePrice As Double
   Dim rowCount As Long
   Dim yearlyChange As Double
   Dim percentChange As Double
   Dim targetRow As Long
   Dim openPriceEachRow As Double
   Dim totalStockVolume As LongLong
   
   For Each ws In Worksheets
   
    Dim bigIncreaseTicker As String
    Dim bigDecreaseTicker As String
    Dim bigIncrease As Double
    Dim bigDecrease As Double
    Dim bigTotalStockVolume As LongLong
    Dim bigTotalStockVolumeTicker As String

   
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
   
    targetRow = 2
    yearlyChange = 0
    percentChange = 0
    totalStockVolume = 0
    rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To rowCount
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            tickerSymbol = ws.Cells(i, 1).Value
            closePrice = ws.Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            percentChange = yearlyChange / openPrice
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
            ws.Cells(targetRow, 9).Value = tickerSymbol
            ws.Cells(targetRow, 10).Value = yearlyChange
            ws.Cells(targetRow, 11).Value = percentChange
            ws.Cells(targetRow, 12).Value = totalStockVolume
            
            ws.Cells(targetRow, 11).NumberFormat = "0.00%"
            
            If yearlyChange > 0 Then
                ws.Cells(targetRow, 10).Interior.ColorIndex = 4
                ws.Cells(targetRow, 11).Interior.ColorIndex = 4
                If bigIncrease < percentChange Then
                    bigIncrease = percentChange
                    bigIncreaseTicker = tickerSymbol
                End If
            
            ElseIf yearlyChange < 0 Then
                ws.Cells(targetRow, 10).Interior.ColorIndex = 3
                ws.Cells(targetRow, 11).Interior.ColorIndex = 3
                If bigDecrease > percentChange Then
                    bigDecrease = percentChange
                    bigDecreaseTicker = tickerSymbol
                End If
            End If
            
            If bigTotalStockVolume < totalStockVolume Then
                bigTotalStockVolume = totalStockVolume
                bigTotalStockVolumeTicker = tickerSymbol
            End If
            
            ' reset before next ticker
            openPrice = ws.Cells(i + 1, 3).Value
            closePrice = 0
            targetRow = targetRow + 1
            percentChange = 0
            yearlyChange = 0
            totalStockVolume = 0
            
        Else
            openPriceEachRow = ws.Cells(i, 3).Value
            
            If (openPrice = 0) And (openPriceEachRow > 0) Then
                openPrice = openPriceEachRow
            End If
            
            totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
            
        End If
        
    Next i
    
    ' add sumary data
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(2, 15).Value = bigIncreaseTicker
    ws.Cells(2, 16).Value = bigIncrease
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(3, 15).Value = bigDecreaseTicker
    ws.Cells(3, 16).Value = bigDecrease
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(4, 15).Value = bigTotalStockVolumeTicker
    ws.Cells(4, 16).Value = bigTotalStockVolume
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"
    
    ' reset before next sheet
    bigIncreaseTicker = ""
    bigDecreaseTicker = ""
    bigTotalStockVolumeTicker = ""
    bigIncrease = 0
    bigDecrease = 0
    bigTotalStockVolume = 0
   Next ws
End Sub
