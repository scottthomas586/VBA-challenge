Attribute VB_Name = "TickerLoop"
Sub TickerSymbol()

'define the variables
For Each ws In Worksheets

    Dim WorksheetName As String
    Dim ticker As String
    Dim volume As LongLong
    Dim summary_ticker As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim i As Long
    WorksheetName = ws.Name
    
    
'set my values
    volume = 0
    summary_ticker = 2
    open_price = ws.Cells(2, 3).Value
    close_price = ws.Cells(2, 6).Value
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'loop through ticker symbols
    For i = 2 To Lastrow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            ws.Range("I" & summary_ticker).Value = ticker
            ws.Range("L" & summary_ticker).Value = volume
            close_price = ws.Cells(i, 6).Value
            yearly_change = (close_price - open_price)
    
                    
            ws.Range("J" & summary_ticker).Value = yearly_change
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If
                
            ws.Range("K" & summary_ticker).Value = percent_change
            ws.Range("K" & summary_ticker).NumberFormat = "0.00%"
            summary_ticker = summary_ticker + 1
            volume = 0
            open_price = Cells(i + 1, 3)
        Else
            volume = volume + Cells(i, 7).Value
            
        End If
        
        
    Next i
    
    'conditional formatting
    
    LastrowTable = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastrowTable
    
        If ws.Cells(i, 10).Value > 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
        
        Else
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
    
        
    Next i
        
    Next ws
        
End Sub






































