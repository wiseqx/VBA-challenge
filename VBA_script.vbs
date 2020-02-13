Sub Stock_market_analyst()
    

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volumn"
        
    
    Dim row_count As Integer
    row_count = 2
    
    Dim Opening_price_row As Double
    Opening_price_row = 2
    
    Dim Total_stock_volume As Double
    Total_stock_volume = 0
    
    num_of_rows = Cells(Rows.Count, 1).End(xlUp).Row
        
    
    For i = 2 To num_of_rows
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ' add unique ticker
            Range("I" & row_count).Value = Cells(i, 1).Value
            
            ' Count yearly change
            Range("J" & row_count).Value = Cells(i, 6).Value - Cells(Opening_price_row, 3).Value
            
            ' Count percent change and use if Statement in case the openning price is equal to 0
            If Cells(Opening_price_row, 3).Value = 0 Then
                Range("K" & row_count).Value = infinity
            Else
                Range("K" & row_count).Value = (Cells(i, 6).Value - Cells(Opening_price_row, 3).Value) / Cells(Opening_price_row, 3).Value
            End If
            
            Opening_price_row = i + 1
            
            ' Count total stock column
            Total_stock_volume = Total_stock_volume + Range("G" & i).Value
            Range("L" & row_count).Value = Total_stock_volume
            
            ' add 1 row to the new table
            row_count = row_count + 1
            
            ' Reset the Total stock volume
            Total_stock_volume = 0
            
        Else
            ' Add to the Total stock volume
            Total_stock_volume = Total_stock_volume + Range("G" & i).Value
        
        End If
        
    Next i
    
    ' highlight positive change in green and negative change in red.
    lastrow = Range("K" & Rows.Count).End(xlUp).Row
    
    For i = 2 To lastrow
        If Range("j" & i).Value > 0 Then
            Range("j" & i).Interior.ColorIndex = 4
        ElseIf Range("j" & i).Value < 0 Then
            Range("j" & i).Interior.ColorIndex = 3
        End If
    Next i
    
    ' Change format
    Range("K2:K" & lastrow).NumberFormat = "0.00%"
        
        
            
End Sub