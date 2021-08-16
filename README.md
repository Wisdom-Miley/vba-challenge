# vba-challenge-homework

Modules 1

```
Sub datastock()

'set worksheet variable and being looping ws
'dim ws As worksheet

For Each ws In Worksheets

Dim Ticker As String

Dim Ticker_Total As Double
Dim Summary_Table_Row As Long


'set Title For Summary Table
ws.Range("I1") = "Ticker Name"
ws.Range("J1") = "Ticker Total"
ws.Range("K1") = "Yearly Change"
ws.Range("L1") = "Percent Change"


'Create Lastrow Variable for Loops

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim openprice As Double
Dim closeprice As Double
Dim pricediff As Double
Dim percentchange As Double
Dim i As Long

openprice = ws.Cells(2, 3).Value

For i = 2 To lastrow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker = ws.Cells(i, 1).Value
Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

ws.Range("I" & Summary_Totle_Row).Value = Ticker
ws.Range("J" & Summary_Totle_Row).Value = Ticker_Total
closeprice = ws.Cells(i, 6).Value
pricediff = closeprice - openprice
ws.Range("K" & Summary_Table_Row).Value = pricediff

If openprice = 0 Then
percentchange = 0

Else
percentchange = pricediff / openprice
End If

ws.Range("L" & Summary_Table_Row).Value = percentchange
ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"

Summary_Table_Row = Summary_Table_Row + 1
Ticker_Total = 0


openprice = ws.Cells(i + 1, 3)

Else
Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

End If
Next i

sumlastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
For j = 2 To latrow
If ws.Cells(j, 11).Value > 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 4

ElseIf ws.Cells(j, 11).Value < 0 Then
ws.Cells(j, 11).Interior.ColorIndex = 3

End If
Next j
Next ws

End Sub
```
