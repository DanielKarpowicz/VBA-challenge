Attribute VB_Name = "Module5"
Sub all_loop()

'loop through all, referenced: https://www.automateexcel.com/vba/cycle-and-update-all-worksheets/
    For Each ws In Worksheets

        'set variables
        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 0

        'location of ticker for summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        'yearly Change is close price - open price
        'percent change is (close - open/open)*100
        Dim open_price As Double
        'initial value
        open_price = ws.Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'headers per homework instructions
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'row count first column; reference: https://www.excelcampus.com/vba/find-last-row-column-cell/
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'set up i loop

        For i = 2 To lastrow

            'detects change on ticker symbol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'set ticker name
              tickername = ws.Cells(i, 1).Value

              'volume of trading
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'ticker name to summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'trade volume for each ticker in the summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'closing price information
              close_price = ws.Cells(i, 6).Value

              'calculate yearly change
               yearly_change = (close_price - open_price)
              
              'yearly change for each ticker in the summary table
              ws.Range("J" & summary_ticker_row).Value = yearly_change

              'make sure to show 0 instead of error
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

            'print the yearly change to the summary table make sure displayed as percentage,
            'reference: https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'reset rown counter and add 1
              summary_ticker_row = summary_ticker_row + 1

              'reset volume back to 0
              tickervolume = 0

              'reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'volume trade addition
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'finding last row of summary table; reference: https://www.excelcampus.com/vba/find-last-row-column-cell/

    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'color coding; use 4 instead of 10 for better green
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

    'labels per homework instructions

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'determine values for the summary table
        For i = 2 To lastrow_summary_table
        
            'max percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'min percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    'need to run, had syntax error
    Next ws
        
End Sub

