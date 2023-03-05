Attribute VB_Name = "Module1"
Sub Module_2()

For Each ws In Worksheets
    
    ws.Range("I1, P1").Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Ticker_Volume = 0
    
    open_count = 0
    
    output_row = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For input_row = 2 To LastRow
    
        If ws.Cells(input_row + 1, 1).Value <> ws.Cells(input_row, 1).Value Then
        
            Ticker_Name = ws.Cells(input_row, 1).Value
            Ticker_Volume = Ticker_Volume + ws.Cells(input_row, 7).Value
        
            close_ = ws.Cells(input_row, 6).Value
            open_ = open_count
            open_val = ws.Cells(input_row - open_, 3).Value
            year_change = close_ - open_val
            percent_change = year_change / open_val
        
            ws.Range("I" & output_row).Value = Ticker_Name
            ws.Range("L" & output_row).Value = Ticker_Volume
            ws.Range("J" & output_row).Value = year_change
            ws.Range("K" & output_row).Value = percent_change
        
            output_row = output_row + 1
        
            Ticker_Volume = 0
            open_count = 0
        
        Else
    
            Ticker_Volume = Ticker_Volume + ws.Cells(input_row, 7).Value
            open_count = open_count + ws.Cells(input_row, 1).Count
        
        End If
        
    
    Next input_row
    
    
    For input_row = 2 To LastRow
    
        If ws.Cells(input_row, 10).Value > 0 Then
        
            ws.Cells(input_row, 10).Interior.ColorIndex = 4
        
        ElseIf ws.Cells(input_row, 10).Value < 0 Then
        
            ws.Cells(input_row, 10).Interior.ColorIndex = 3
        
        Else
        
        End If
    
    Next input_row
    
    
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    MaxValue = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 17).Value = MaxValue
    
    MinValue = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 17).Value = MinValue
    
    MaxValueVol = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 17).Value = MaxValueVol
    
    For input_row = 2 To LastRow
        If ws.Cells(input_row, 11).Value = MaxValue Then
        
            ws.Cells(2, 16).Value = ws.Cells(input_row, 9).Value
            
        End If
        
        If ws.Cells(input_row, 11).Value = MinValue Then
        
            ws.Cells(3, 16).Value = ws.Cells(input_row, 9).Value
            
        End If
        
        If ws.Cells(input_row, 12).Value = MaxValueVol Then
        
            ws.Cells(4, 16).Value = ws.Cells(input_row, 9).Value
            
        End If
        
    Next input_row
    
    ws.Range("J:J").NumberFormat = "0.00"
    
    ws.Range("K:K, Q2:Q3").NumberFormat = "0.00%"
    
    ws.Columns("I:Q").AutoFit

Next ws
    
End Sub
