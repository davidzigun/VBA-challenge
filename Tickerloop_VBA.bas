Attribute VB_Name = "Module1"

Sub tickerloop()

    'Variable for ticker name
        Dim ticker As String
    'Setting variable for open price
        Dim open_price As Double
        
    
    'Setting variables for finding difference and change in price
        Dim close_price As Double
        Dim yearly_range As Double
        Dim percent_change As Double
    
    'Total count of total trade volume
        Dim tickervolume As Double
        tickervolume = 0

        'Track each ticker name in the summary table
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2

        'Summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Count of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through rows by the ticker names
   
        For i = 2 To lastrow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tickername = Cells(i, 1).Value
                tickervolume = tickervolume + Cells(i, 7).Value
                Range("I" & summary_ticker_row).Value = tickername
                Range("L" & summary_ticker_row).Value = tickervolume
  
                close_price = Cells(i, 6).Value
                yearly_change = (close_price - open_price)
                 Range("J" & summary_ticker_row).Value = yearly_change
                
                    If (open_price = 0) Then
                        percent_change = 0
                    
                    Else
                        percent_change = yearly_change / open_price
                        
                    End If
              
              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
                
                'Add one to the summary_ticker_row
                summary_ticker_row = summary_ticker_row + 1

                'Reset tickervolume to zero
                tickervolume = 0
                
                'Reset the opening price
                open_price = Cells(i + 1, 3)

            Else
              
                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

            End If
        
        Next i

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code change
    
    For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

        'Defining headers for Ticker and percent change
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
    
    'Row labels for Highest and lowest change
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"

    'Finding Max and Min in Percentage change column and Max Volum in Volume Total column
        
        For i = 2 To lastrow_summary_table
            
            'Maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Max volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i

End Sub
