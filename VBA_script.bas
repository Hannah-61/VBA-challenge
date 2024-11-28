Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim rowOutput As Integer
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestTickerIncrease As String
    Dim greatestTickerDecrease As String
    Dim greatestTickerVolume As String

    ' Loop through all worksheets
    For Each ws In Worksheets
        ws.Activate
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables
        rowOutput = 2
        totalVolume = 0
        openPrice = ws.Cells(2, 3).Value ' First open price
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Output Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 15).Value = "Greatest % Increase"
        ws.Cells(1, 16).Value = "Greatest % Decrease"
        ws.Cells(1, 17).Value = "Greatest Total Volume"
        
        ' Loop through rows
        For i = 2 To lastRow
            ' Add volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if the ticker changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                
                ' Output data
                ws.Cells(rowOutput, 9).Value = ticker
                ws.Cells(rowOutput, 10).Value = totalVolume
                ws.Cells(rowOutput, 11).Value = closePrice - openPrice
                
                If openPrice <> 0 Then
                    ws.Cells(rowOutput, 12).Value = ((closePrice - openPrice) / openPrice) * 100
                Else
                    ws.Cells(rowOutput, 12).Value = 0
                End If
                
                ' Check for greatest % increase/decrease and volume
                If ws.Cells(rowOutput, 12).Value > greatestIncrease Then
                    greatestIncrease = ws.Cells(rowOutput, 12).Value
                    greatestTickerIncrease = ticker
                End If
                
                If ws.Cells(rowOutput, 12).Value < greatestDecrease Then
                    greatestDecrease = ws.Cells(rowOutput, 12).Value
                    greatestTickerDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestTickerVolume = ticker
                End If
                
                ' Reset variables
                rowOutput = rowOutput + 1
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            End If
        Next i

        ' Output greatest values
        ws.Cells(2, 15).Value = greatestTickerIncrease & " (" & greatestIncrease & "%)"
        ws.Cells(2, 16).Value = greatestTickerDecrease & " (" & greatestDecrease & "%)"
        ws.Cells(2, 17).Value = greatestTickerVolume & " (" & greatestVolume & ")"

        ' Apply conditional formatting
        With ws.Range("K2:K" & rowOutput - 1).FormatConditions.Add(xlCellValue, xlGreater, "0")
            .Interior.Color = RGB(0, 255, 0)
        End With
        With ws.Range("K2:K" & rowOutput - 1).FormatConditions.Add(xlCellValue, xlLess, "0")
            .Interior.Color = RGB(255, 0, 0)
        End With
    Next ws
End Sub

