Sub Worksheet()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate

    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest Total Volume "
    Columns("N").AutoFit

    Call StockInfo
    Call Summary
    Call HeadersStock

Next

End Sub


Sub StockInfo()
Dim column As Integer
Dim lastrow As Long
Dim ticker_name As String
Dim total, row, iter As Long
Dim year_change, per_change As Double
Dim open_price, close_price As Double

open_price = 0
close_price = 0
year_change = 0
per_change = 0
total = 0
column = 1
lastrow = Cells(Rows.Count, 1).End(xlUp).row
row = 2
iter = 0

For i = 2 To lastrow

    If Cells(i, column).Value <> Cells(i + 1, column).Value Then
    
        ticker_name = Cells(i, 1).Value
        Range("I" & row).Value = ticker_name
        
        addition = Cells(i, 7).Value + Cells(i + 1, 7).Value
        total = total + Cells(i, 7).Value
        Range("L" & row).Value = total
        
        close_price = Cells(i, 6).Value
        open_price = Cells(i - iter, 3).Value
        year_change = close_price - open_price
        Range("J" & row).Value = year_change
        
        If year_change < 0 Then
            Range("J" & row).Interior.ColorIndex = 3
        Else
            Range("J" & row).Interior.ColorIndex = 4
        End If
    
        If open_price = 0 Then
        
        Else
            per_change = (year_change / open_price) * 100
            Range("K" & row).Value = per_change
        End If
        
        row = row + 1
        open_price = 0
        close_price = 0
        year_change = 0
        per_change = 0
        iter = 0
        total = 0
        
    Else
        total = total + Cells(i, 7).Value
        iter = iter + 1
        
    End If
    
Next i

End Sub


Sub Summary()

Dim finalrow, volume, x, y As Long
Dim increase, decrease As Double
Dim ticker1, ticker2, ticker3 As String
Dim z As Integer

increase = 0
decrease = 0
volume = 0
finalrow = Cells(Rows.Count, 1).End(xlUp).row

For x = 2 To finalrow
    If Cells(x, 11).Value > increase Then
        increase = Cells(x, 11).Value
    Else
        increase = increase
    End If
    If Cells(x, 11).Value < decrease Then
        decrease = Cells(x, 11).Value
    Else
        decrease = decrease
    End If
     If Cells(x, 12).Value > volume Then
        volume = Cells(x, 12).Value
    Else
        volume = volume
    End If
Next x

Range("P2").Value = increase
Range("P3").Value = decrease
Range("P4").Value = volume

For y = 2 To finalrow

   If Cells(y, 11).Value = Range("P2").Value Then
        ticker1 = Cells(y, 9).Value
        Range("O2").Value = ticker1
   End If
   
    If Cells(y, 11).Value = Range("P3").Value Then
        ticker2 = Cells(y, 9).Value
         Range("O3").Value = ticker2
    End If
   
    If Cells(y, 12).Value = Range("P4").Value Then
        ticker3 = Cells(y, 9).Value
        Range("O4").Value = ticker3
   End If
   
Next y

End Sub

Sub HeadersStock()

Dim Headers
Dim i, j As Integer

Headers = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
j = 9

For i = LBound(Headers) To UBound(Headers)
    Cells(1, j).Value = Headers(i)
    j = j + 1
Next i

Columns("J").AutoFit
Columns("K").AutoFit
Columns("L").AutoFit

End Sub
