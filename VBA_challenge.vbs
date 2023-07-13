Sub ticker_symbol()

For Each ws In Worksheets
    
    Dim Symbol As String
    Dim ticker_volume As Double
            ticker_volume = 0
    Dim Ticker_Table As Long
    Ticker_Table = 1
    Dim open_price As Double
    Dim close_price  As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim last_row As Long
     
    
        
    

    ws.Activate
    ws.Cells(1, 1).Activate
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For i = 2 To last_row
        'Last row of group'
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Symbol = ws.Cells(i, 1).Value
            Ticker_Table = Ticker_Table + 1
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            ws.Range("I" & Ticker_Table).Value = Symbol
            ws.Range("L" & Ticker_Table).Value = ticker_volume
            ticker_volume = 0
            close_price = ws.Cells(i, 6).Value
            yearly_change = close_price - open_price
            percent_change = (close_price - open_price) / open_price
            ws.Range("j" & Ticker_Table).Value = yearly_change
            ws.Range("k" & Ticker_Table).Value = percent_change
            
       
        'first row of group'
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            open_price = ws.Cells(i, 3).Value
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            
        
        
         'middle row of group'
         
         Else
            ticker_volume = ticker_volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i

Next ws


End Sub
