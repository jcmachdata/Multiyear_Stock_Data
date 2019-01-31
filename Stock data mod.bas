Attribute VB_Name = "Module1"
Sub stockdata()

Dim ws As Worksheet

For Each ws In Worksheets

Dim total As Double
Dim j As Long
Dim lastrow As Long
Dim i As Long
Dim openstock As Double
Dim closestock As Double


' ignore headers
j = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

' set initial ticker value and openstock value
ws.Cells(2, 9).Value = ws.Cells(2, 1).Value
openstock = ws.Cells(2, 3).Value

' iterate over each ticker and add the volume together
For i = 2 To (lastrow + 1)
    
    If ws.Cells(i, 1).Value = ws.Cells(j, 9).Value Then

        total = total + ws.Cells(i, 7).Value
    
    ' if new ticker encountered, output the old symbol and total volume
    Else
        closestock = ws.Cells(i - 1, 6).Value
        ws.Cells(j, 10).Value = closestock - openstock
        
            ' Set interior color for yearly change
            If ws.Cells(j, 10) >= 0 Then
                 ws.Cells(j, 10).Interior.ColorIndex = 4
            Else: ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            ' Handle special cases of opening and closing stock value
            If openstock = 0 And closestock <> 0 Then
        
                ws.Cells(j, 11).Value = 1
            
            ElseIf openstock = 0 And closestock = 0 Then
                ws.Cells(j, 11).Value = 0
            
            ' calculate percent change
            Else: ws.Cells(j, 11).Value = (closestock - openstock) / openstock
                  ws.Cells(j, 11).NumberFormat = "0.00%"
        
            End If
        
        ' reset openstock price, output total volume then reset total volume
        openstock = ws.Cells(i, 3).Value
        ws.Cells(j, 12).Value = total
        total = ws.Cells(i, 7).Value
        j = j + 1
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
  
    
    End If


Next i

Columns("A:L").AutoFit

Next ws

End Sub
