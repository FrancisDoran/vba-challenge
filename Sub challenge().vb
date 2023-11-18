Sub challengeAllSheets()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        Call challenge(ws)
    Next ws
End Sub

Sub challenge(ws As Worksheet)

   
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As Double
    Dim startPrice As Double
    Dim endPrice As Double
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    j = 2 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value= "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            endPrice = ws.Cells(i, 6).Value 
            yearlyChange = endPrice - startPrice
            If startPrice <> 0 Then
                percentChange = yearlyChange / startPrice
            Else
                percentChange = 0
            End If
            totalStockVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7))) 
            
            ws.Cells(j, 9).Value = ticker
            ws.Cells(j, 10).Value = yearlyChange
            ws.Cells(j, 11).Value = percentChange
            ws.Cells(j, 12).Value = totalStockVolume
            
            j = j + 1
            startPrice = 0
            totalStockVolume = 0
        ElseIf startPrice = 0 Then
            startPrice = ws.Cells(i, 3).Value 
        End If
    Next i
    For i = 2 To j
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0) 
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        End If
    Next i
    Dim max As Double
    Dim min As Double
    Dim vol_max As Double
    ' find values
    max = Application.WorksheetFunction.Max(ws.Range("K1:K" & lastRow))
    min = Application.WorksheetFunction.Min(ws.Range("K1:K" & lastRow))
    vol_max = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    ' set values
    ws.Cells(2, 17).Value = max
    ws.Cells(3, 17).Value = min
    ws.Cells(4, 17).Value = vol_max
    ' find ticker for values
    max_ticker = Application.WorksheetFunction.Match(max, ws.Range("K2:K" & lastRow), 0) + 1
    min_ticker = Application.WorksheetFunction.Match(min, ws.Range("K2:K" & lastRow), 0) + 1
    vol_ticker = Application.WorksheetFunction.Match(vol_max, ws.Range("L2:L" & lastRow), 0) + 1
    ' set values
    ws.Cells(2, 16).Value = ws.Cells(max_ticker,9)
    ws.Cells(3, 16).Value = ws.Cells(min_ticker,9)
    ws.Cells(4, 16).Value = ws.Cells(vol_ticker,9)
End Sub 
