Attribute VB_Name = "Module1"
Sub stockdata()

Dim ws As Worksheet

For Each ws In Worksheets

Dim total As Double
Dim j As Long
Dim lastrow As Long
Dim i As Long

' ignore headers
j = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Total Stock Volume"

' set initial ticker value
ws.Cells(2, 9).Value = ws.Cells(2, 1).Value

' iterate over each ticker and add the volume together
For i = 2 To (lastrow + 1)

    If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then

        total = total + ws.Cells(i, 7).Value
    
        ' if new ticker encountered, output the old symbol and total volume
    Else
    
        ws.Cells(j, 10).Value = total
        total = ws.Cells(i, 7).Value
        j = j + 1
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
  
    
    End If


Next i

Columns("A:J").AutoFit

Next ws

End Sub
