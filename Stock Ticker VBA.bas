Attribute VB_Name = "Module1"
Sub TickerSymbol()
Dim ticker As String
Dim volume As Double
volume = 0
Dim summary_ticker As Integer
summery_ticker = 2
Dim open_price As Double
open_price = Cells(2, 3).Value
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'I can't seem to get the LastRow formula to work.  It keeps giving me an error message.  I literally tried for 5 hours to fix.
'LastRow = Cells(Rows.count, 1).End(x1Up).row

For i = 2 To 70926

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
        Range("I" & summary_ticker).Value = ticker
        Range("L" & summary_ticker).Value = volume
        close_price = Cells(i, 6).Value
        yearly_change = (close_price - open_price)
        Range("J" & summary_ticker).Value = yearly_change
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / open_price
            End If
        
        Range("K" & summary_ticker).Value = percent_change
        Range("K" & summary_ticker).NumberFormat = "o.00%"
        summary_ticker = summary_ticker + 1
        volume = 0
        open_price = Cells(i + 1, 3)
    Else
        volume = volume + Cells(i, 7).Value
        
    End If

Next i

End Sub
