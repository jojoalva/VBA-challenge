Attribute VB_Name = "Module2"
Sub smalltable()

Dim maxticker As String
Dim maxincrease As Double
Dim minticker As String
Dim minincrease As Double
Dim greatstockvol As LongLong
Dim greatticker As String
Dim klastrow As Long

For Each ws In Worksheets


ws.Range("O1").Value = "Ticker"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Stock Volume"

greatstockvol = 0
maxincrease = 0
minincrease = 0
klastrow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

'find most positive % increase & corresponding ticker
For i = 2 To klastrow

If ws.Cells(i, 11).Value > maxincrease Then
                maxincrease = ws.Cells(i, 11).Value
                maxticker = ws.Cells(i, 9).Value
End If
'print these values in cells
ws.Range("O2").Value = maxincrease
ws.Range("P2").Value = maxticker
                
'find most negative % decrease & corresponding ticker
If ws.Cells(i, 11).Value < minincrease Then
                minincrease = ws.Cells(i, 11).Value
                minticker = ws.Cells(i, 9).Value
                
End If
'print these values in cells
ws.Range("O3").Value = minincrease
ws.Range("P3").Value = minticker
                
'find greatest totalstock volume & corresponding ticker
If ws.Cells(i, 12).Value > greatstockvol Then
greatstockvol = ws.Cells(i, 12).Value
greatticker = ws.Cells(i, 9).Value
                
End If

'print these values in cells
ws.Range("O4").Value = greatstockvol
ws.Range("P4").Value = greatticker

'set maxincrease to percent format
ws.Range("O2:O3").Style = "percent"

Next i

'autofit table headers
ws.Columns("N:P").AutoFit

Next ws

End Sub

