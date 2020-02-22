Sub StockCalculations()

    For Each ws In Worksheets
        
        Dim ticker As String
        Dim dollar_change As Double
        Dim percent_change As Double
        Dim total_volume As Double
    
        Dim opening_price As Double
        Dim closing_price As Double
        
        Dim summary_table_row As Long
        summary_table_row = 2
    
        'total_volume start number
        total_volume = 0
    
        'last row
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Add Titles to All Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
    
    
        'Creating Summary Table
    
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closing_price = ws.Cells(i, 6).Value
                dollar_change = closing_price - opening_price
                
                If opening_price <> 0 Then
                    percent_change = (closing_price - opening_price) / opening_price
                Else
                    percent_change = 0
                End If
                
                total_volume = total_volume + ws.Cells(i, 7).Value
            
                ws.Range("I" & summary_table_row).Value = ticker
                ws.Range("J" & summary_table_row).Value = dollar_change
                ws.Range("K" & summary_table_row).Value = percent_change
                ws.Range("L" & summary_table_row).Value = total_volume
            
                If ws.Range("J" & summary_table_row).Value > 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
            
                summary_table_row = summary_table_row + 1
                total_volume = 0
                opening_price = 0
                closing_price = 0
        
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                opening_price = ws.Cells(i, 3).Value
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
            
            End If
        
            ws.Cells(i, 10).NumberFormat = "$#,##0.000000000"
            ws.Cells(i, 11).NumberFormat = "0.00%"
        
        Next i
    
    
    'Max% Min% and Greatest Total Volume
        
        Dim lastrow_summary_table_2 As Long
        lastrow_summary_table_2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        Dim max_percent_change As Double
        max_percent_change = 0
        Dim max_percent_change_ticker As String
        
        Dim min_percent_change As Double
        min_percent_change = 0
        Dim min_percent_change_ticker As String
    
        Dim greatest_total_volume As Double
        greatest_total_volume = 0
        Dim greatest_total_volume_ticker As String
    
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
    
        For j = 2 To lastrow_summary_table_2
        
            If ws.Cells(j, 11).Value > max_percent_change Then
                max_percent_change = ws.Cells(j, 11)
                max_percent_change_ticker = ws.Cells(j, 9)
                ws.Range("P2").Value = max_percent_change
                ws.Range("O2").Value = max_percent_change_ticker
        
            End If
        
            If ws.Cells(j, 11).Value < min_percent_change Then
                min_percent_change = ws.Cells(j, 11)
                min_percent_change_ticker = ws.Cells(j, 9)
                ws.Range("P3").Value = min_percent_change
                ws.Range("O3").Value = min_percent_change_ticker
        
            End If
        
            If ws.Cells(j, 12).Value > greatest_total_volume Then
                greatest_total_volume = ws.Cells(j, 12)
                greatest_total_volume_ticker = ws.Cells(j, 9)
                ws.Range("P4").Value = greatest_total_volume
                ws.Range("O4").Value = greatest_total_volume_ticker
                
            End If
          
        Next j
        
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "General"
    
    Next ws
   
End Sub
