Attribute VB_Name = "Module2"
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
