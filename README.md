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

Module 2


```
Sub Challege()

   
    Dim i As Long
    Dim j As Integer
    Dim TopDog As Double
    Dim UnderDog As Double
    Dim RowNum As Long
   
    Dim ws As Worksheet
    
    Dim HiTicker As String
    Dim LwTicker As String
  
    Dim PstOS As Integer

    ' Loop over worksheets
    For Each ws In Worksheets
        ' Title headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(5, 14).Value = "Least Total Volume"
        
        ' Set paste offset, highest number, lowest number variables to zero
        PstOS = 0
        TopDog = 0
        UnderDog = 0
        ' Determine number of rows in sheet
        RowNum = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Loop over columns
        For j = 11 To 12
            ' Loop over rows
            For i = 2 To RowNum
                
                If ws.Cells(i, j).Value > TopDog Then
                    TopDog = ws.Cells(i, j).Value
                    HiTicker = ws.Cells(i, 9).Value
                Else
                End If
                
                If ws.Cells(i, j).Value < UnderDog Then
                    UnderDog = ws.Cells(i, j).Value
                    LwTicker = ws.Cells(i, 9).Value
                Else
                End If
                
            Next i

        ws.Cells(2 + PstOS, 15).Value = HiTicker
        ws.Cells(2 + PstOS, 16).Value = TopDog
        ' Same with underdog
        ws.Cells(3 + PstOS, 15).Value = LwTicker
        ws.Cells(3 + PstOS, 16).Value = UnderDog
        ' Reset all variables
        HiTicker = "None"
        LwTicker = "None"
        TopDog = 0
        UnderDog = 0
     
        PstOS = PstOS + 2
        Next j
   
    ws.Range("N5:P5").Value = ""
    ws.Range("P2:P3").NumberFormat = "0.00%"
    Next ws
End Sub
```
