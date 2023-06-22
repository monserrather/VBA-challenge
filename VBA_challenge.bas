Attribute VB_Name = "Module1"
Sub Button1_Click()
'Loops through all the sheets
For Each ws In Worksheets

    'Creation of columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    'This block of code gets all the distinct tickers
    
    Dim v_control As String
    Dim TotalRows As Long
    Dim total_ticker As Double
    Dim price_end As Double
    Dim price_start As Double
    Dim row_Result As Integer
    
    v_control = ws.Cells(2, 1).Value
    TotalRows = ws.Rows(Rows.Count).End(xlUp).Row
    total_ticker = 0
    price_start = ws.Cells(2, 3).Value
    row_Result = 2
    
    For i = 1 To TotalRows
        If (v_control = ws.Cells(i + 1, 1)) Then
            'Acumulate number of stocks
            total_ticker = total_ticker + CDbl(ws.Cells(i + 1, 7).Value)
            
        Else
            price_end = ws.Cells(i, 6).Value
            'Populate stock name
            ws.Cells(row_Result, 9).Value = v_control
            'Populate total stocks
            ws.Cells(row_Result, 12).Value = total_ticker
            'Populate yearly change
            ws.Cells(row_Result, 10).Value = price_end - price_start
            'Populate percentage
            ws.Cells(row_Result, 11).Value = FormatPercent(ws.Cells(row_Result, 10).Value / price_start, 2)
            'Change color based on condition
            If (ws.Cells(row_Result, 10).Value < 0) Then
                ws.Cells(row_Result, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(row_Result, 10).Interior.ColorIndex = 10
            End If
            
            'Variable control changes
            v_control = ws.Cells(i + 1, 1)
            'Total ticker is reinitialized
            total_ticker = CDbl(ws.Cells(i + 1, 7).Value)
            'price_start is set
            price_start = ws.Cells(i + 1, 3).Value
            row_Result = row_Result + 1
        End If
    Next i

 'This iterates over all calculated values and extracts min and max
    
    'Declare and initialize accumulator variables
    Dim gr_per As Double
    Dim lw_per As Double
    Dim gr_tot As Double
    
    row_Result = row_Result - 2
    
    gr_per = ws.Cells(2, 11).Value
    lw_per = ws.Cells(2, 11).Value
    gr_tot = ws.Cells(2, 12).Value
    gr_per_ticker = ws.Cells(2, 9).Value
    lw_per_ticker = ws.Cells(2, 9).Value
    gr_tot_ticker = ws.Cells(2, 9).Value
    
    For count_final_values = 1 To row_Result
        
        'This compares the values to get the max %
        If (ws.Cells(count_final_values + 1, 11).Value > gr_per) Then
            gr_per = ws.Cells(count_final_values + 1, 11).Value
            gr_per_ticker = ws.Cells(count_final_values + 1, 9).Value
        End If
        
        'This compares the values to get the min %
        If (ws.Cells(count_final_values + 1, 11).Value < lw_per) Then
            lw_per = ws.Cells(count_final_values + 1, 11).Value
            lw_per_ticker = ws.Cells(count_final_values + 1, 9).Value
        End If
        
        'This compares the values to get the max total
        If (ws.Cells(count_final_values + 1, 12).Value > gr_tot) Then
            gr_tot = ws.Cells(count_final_values + 1, 12).Value
            gr_tot_ticker = ws.Cells(count_final_values + 1, 9).Value
        End If
            
    Next count_final_values
    
    'Print greatest % increase
    ws.Cells(2, 16).Value = gr_per_ticker
    ws.Cells(2, 17).Value = FormatPercent(gr_per, 2)
    
    'Print lowest % increase
    ws.Cells(3, 16).Value = lw_per_ticker
    ws.Cells(3, 17).Value = FormatPercent(lw_per, 2)
    
    'Print greatest total
    ws.Cells(4, 16).Value = gr_tot_ticker
    ws.Cells(4, 17).Value = gr_tot
   

Next ws
End Sub
