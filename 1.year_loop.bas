
Sub year_loop()

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
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'headers per homework instructions
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'row count first column; reference: https://www.excelcampus.com/vba/find-last-row-column-cell/
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'set up i loop

        For i = 2 To lastrow

            'detects change on ticker symbol
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Set ticker name
              tickername = Cells(i, 1).Value

              'volume of trading
              tickervolume = tickervolume + Cells(i, 7).Value

              'ticker name to summary table
              Range("I" & summary_ticker_row).Value = tickername

              'trade volume for each ticker in the summary table
              Range("L" & summary_ticker_row).Value = tickervolume

              'closing price information
              close_price = Cells(i, 6).Value

              'calculate yearly change
              yearly_change = (close_price - open_price)
              
              'yearly change for each ticker in the summary table
              Range("J" & summary_ticker_row).Value = yearly_change

             'make sure to show 0 instead of error
                If (open_price = 0) Then
                    percent_change = 0

                Else
                    percent_change = yearly_change / open_price
                
                End If
            'print the yearly change to the summary table make sure displayed as percentage,
            'reference: https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'reset rown counter and add 1
              summary_ticker_row = summary_ticker_row + 1

              'reset volume back to 0
              tickervolume = 0

              'reset the opening price
              open_price = Cells(i + 1, 3)
            
            Else
              
               'volume trade addition
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i

    'finding last row of summary table; reference: https://www.excelcampus.com/vba/find-last-row-column-cell/

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'color coding; use 4 instead of 10 for better green
    
    For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

End Sub
