# VBA_Challenge
Sub Module2Challenge()

Dim ws As Worksheet

For Each ws In Worksheets
    
    'Initial variable for holding the ticker name, yearly change, percent change and total stock volume
    Dim ticker_name As String
    Dim open_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As LongLong
    Dim max_volume As LongLong
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    volume = 0
    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Keep track of the location for each ticker name in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    Start = 2
    total_stock_volume = 0
    
    'Add column headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'add row headers
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'loop through all tickers
    For i = 2 To lastrow
        
        'check if we are still within the same ticker name, if it is not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'set the ticker name and add it to the summary table
            ticker_name = ws.Cells(i, 1).Value
            ws.Cells(summary_table_row, 9).Value = ticker_name
            Opening_price = ws.Cells(Start, 3).Value
            closing_price = ws.Cells(i, 6).Value
            
            'Set yearly change total and add it to the summary table
            yearly_change = (closing_price - Opening_price)
            
            ws.Cells(summary_table_row, 10).Value = yearly_change
                
                'fix error
                If (Opening_price = 0) Then
                percent_change = 0
                Else
                
                'set Percent change, add it to the summary table and format
                percent_change = (yearly_change / ws.Cells(Start, 3)) * 1
                Start = i + 1
                ws.Cells(summary_table_row, 11).Value = percent_change
                ws.Cells(summary_table_row, 11).NumberFormat = "0.00%"
                End If
            
            'set total stock volume and add to summary table
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            ws.Cells(summary_table_row, 12).Value = total_stock_volume
         
         'add to the summary table row
            summary_table_row = summary_table_row + 1
        
        'reset values
            total_stock_volume = 0
            Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        End If
        Next i
  
  For i = 2 To lastrow_summary_table
            
            'Check if yearly change is + or -
            If ws.Cells(i, 10).Value >= 0 Then
            
            'color + value green
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
            'color - value red
            Else: ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i
        
        'find max increase and decrease, and greatest volume
        Max_Increase = WorksheetFunction.Max(ws.Columns("K"))
        Max_Decrease = WorksheetFunction.Min(ws.Columns("K"))
        max_volume = WorksheetFunction.Max(ws.Columns("L"))
        
        For i = 2 To lastrow_summary_table
        
        If ws.Cells(i, 11) = Max_Increase Then
        ws.Cells(2, 16).Value = Cells(i, 9).Value
        ws.Cells(2, 17).Value = FormatPercent(Max_Increase)
        End If
        
        If ws.Cells(i, 11) = Max_Decrease Then
        ws.Cells(3, 16).Value = Cells(i, 9).Value
        ws.Cells(3, 17).Value = FormatPercent(Max_Decrease)
        End If
        
        If ws.Cells(i, 12) = max_volume Then
        ws.Cells(4, 16).Value = Cells(i, 9).Value
        ws.Cells(4, 17).Value = max_volume
        End If
        
       Next i
 Next ws
End Sub
