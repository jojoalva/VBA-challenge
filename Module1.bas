Attribute VB_Name = "Module1"
Sub alphabetticker()
Dim ws As Worksheet
Dim tickername As String
Dim summaryrow As Long
Dim lastrow As Long
Dim totalstockvolume As LongLong
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim tickerrow As Long


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker Name"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

summaryrow = 2
tickerrow = 2
lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    ' Check if the next ticker cell is the same as the current ticker cell, if not ...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' set opening price of stock on first working day of year
         openprice = ws.Cells(tickerrow, 3).Value
         
        'set closing price of stock on last working day of year
        closeprice = ws.Cells(i, 6).Value
        
        'set yearlychange
        yearlychange = closeprice - openprice
        
        'apply conditional formatting to yearlychange
        If yearlychange >= 0 Then
        ws.Range("J" & summaryrow).Interior.ColorIndex = 4
        Else
        ws.Range("J" & summaryrow).Interior.ColorIndex = 3
        End If
               
        'print in summary table yearlychange
        ws.Range("J" & summaryrow).Value = yearlychange
               
        'work out % change
        percentchange = yearlychange / openprice
        
        'print the percent change
        ws.Range("K" & summaryrow).Value = percentchange
                               
        'reset tickerrow
        tickerrow = i + 1
        
       ' set tickername value
        tickername = ws.Cells(i, 1).Value
        
       'set total stock value for each tickername
        totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
       
       ' print tickername value into summary row cell
        ws.Range("I" & summaryrow).Value = tickername
        
       'print total stock volume for tickername
        ws.Range("L" & summaryrow).Value = totalstockvolume
               
        'go to next summary row
        summaryrow = summaryrow + 1
        totalstockvolume = 0

 
    Else
           
       'if tickername is the same
       totalstockvolume = totalstockvolume + ws.Cells(i, 7).Value
       
    End If

Next i

'autofit summary and last table columns
ws.Columns("I:L").AutoFit
    
'format to % sign
ws.Range("K:K").NumberFormat = "#0.00%"

Next ws

End Sub

