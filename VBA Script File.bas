Attribute VB_Name = "Module1"
Sub AlphabeticalTesting()


Dim ws As Worksheet

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

SummaryRow = 2

For i = 2 To LastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value

OpeningPrice = ws.Cells(i, 3).Value
ClosingPrice = ws.Cells(i, 6).Value
YearlyChange = ClosingPrice - OpeningPrice

If OpeningPrice <> 0 Then
PercentChange = (YearlyChange / OpeningPrice) * 100
Else
PercentChange = 0
End If


ws.Cells(SummaryRow, 9).Value = Ticker
ws.Cells(SummaryRow, 10).Value = YearlyChange
ws.Cells(SummaryRow, 11).Value = PercentChange
ws.Cells(SummaryRow, 12).Value = TotalStockVolume

Cells(SummaryRow, 11).NumberFormat = "0.00%"

If YearlyChange > 0 Then
ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
ElseIf YearlyChange < 0 Then
ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
End If

SummaryRow = SummaryRow + 1

TotalVolume = 0

End If

TotalVolume = TotalVolume + Cells(i, 7).Value
Next i

Next ws

End Sub
