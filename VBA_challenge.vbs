Sub Summary_Table()
    
    For Each ws In Worksheets

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volumn"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        lastrow = ws.Range("K" & Rows.Count).End(xlUp).Row
        
        ' find the maximum and the minimum number
        max_num = 0
        min_num = 0
        vol_num = 0
        For i = 2 To lastrow
            If ws.Range("K" & i).Value > max_num Then
                max_num = ws.Range("K" & i).Value
                max_ticker_row = i
                
            ElseIf ws.Range("K" & i).Value < min_num Then
                min_num = ws.Range("K" & i).Value
                min_ticker_row = i
                
            ElseIf ws.Range("L" & i).Value > vol_num Then
                vol_num = ws.Range("L" & i).Value
                vol_ticker_row = i
            End If
            
        Next i
        
        ws.Range("P2").Value = ws.Range("I" & max_ticker_row).Value
        ws.Range("Q2").Value = max_num
        
        ws.Range("P3").Value = ws.Range("I" & min_ticker_row).Value
        ws.Range("Q3").Value = min_num
        
        ws.Range("P4").Value = ws.Range("I" & vol_ticker_row).Value
        ws.Range("Q4").Value = vol_num
        
        ' Formatting
        ws.Range("Q2:Q3" & new_lastrow).NumberFormat = "0.00%"

        
    Next ws
    
            
End Sub
