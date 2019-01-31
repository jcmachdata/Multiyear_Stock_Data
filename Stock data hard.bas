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
Dim k As Long
Dim gpercent As Double
Dim m As Long
Dim gpercentdec As Double
Dim gvolume As Double
Dim n As Long


' ignore headers
j = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1,P1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("Q1") = "Value"

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

' set initial value of ticker with greatest percent increase
ws.Cells(2, 16).Value = ws.Cells(2, 9).Value
ws.Cells(2, 17).Value = ws.Cells(2, 11).Value
gpercent = ws.Cells(2, 11).Value

    ' find ticker with greatest percent increase
    For k = 3 To lastrow
    
        If ws.Cells(k, 11).Value > gpercent Then
            ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
            gpercent = ws.Cells(k, 11).Value

        End If
        
    Next k

ws.Cells(2, 17).NumberFormat = "0.00%"

' set initial value of ticker with greatest percent decrease
ws.Cells(3, 16).Value = ws.Cells(2, 9).Value
ws.Cells(3, 17).Value = ws.Cells(2, 11).Value
gpercentdec = ws.Cells(2, 11).Value

    ' find ticker with greatest percent decrease
    For m = 3 To lastrow
        
        If ws.Cells(m, 11).Value < gpercentdec Then
            ws.Cells(3, 16).Value = ws.Cells(m, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(m, 11).Value
            gpercentdec = ws.Cells(m, 11).Value

        End If
    
    Next m
    
ws.Cells(3, 17).NumberFormat = "0.00%"

' set initial value of ticker with greatest volume
ws.Cells(4, 16).Value = ws.Cells(2, 9).Value
ws.Cells(4, 17).Value = ws.Cells(2, 12).Value
gvolume = ws.Cells(2, 12).Value

    ' find ticker with greatest volume
    For n = 3 To lastrow
        
        If ws.Cells(n, 12).Value > gvolume Then
            ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(n, 12).Value
            gvolume = ws.Cells(n, 12).Value

        End If
    
    Next n
    

Columns("A:Q").AutoFit

Next ws

End Sub
