# VBA-Script-Challenge-2
I attached the excel file and screen shots along with the VBA code in a seperate file and on the submission. 
Below is code I wrotes for the challenge. 


Sub stocks()

'loop all sheets
    For Each ws In Worksheets

       
       'Set a variable for holding the ticker name, the column of interest
        Dim tickername As String
    
        'Set a varable for holding a total count on the total volume of trade
        Dim tickervolume As Double
        tickervolume = 0
        'add location of ticker count
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        
        Dim open_price As Double
        'pull first price per ticker
        open_price = ws.Cells(2, 3).Value
        'set variables for rest of data points
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Label summary tables
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column.
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       'start inner loop to pull data
     
   For i = 2 To LastRow

            'Searches for when the value of the next ticker is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              tickername = ws.Cells(i, 1).Value

              'sum volume of trades
              tickervolume = tickervolume + ws.Cells(i, 7).Value

              'put tickername in summary table
              ws.Range("I" & summary_ticker_row).Value = tickername

              'put ticker total volume in summary table
              ws.Range("L" & summary_ticker_row).Value = tickervolume

              'Pull close price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'put yearly change in summary table
              ws.Range("J" & summary_ticker_row).Value = yearly_change

              'Put 0 if there is no open price
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              'put yearly change into summary table
              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the row counter. and go to next line on summary table
              summary_ticker_row = summary_ticker_row + 1

              'Reset volume
              tickervolume = 0

              'Go to next open price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
              tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

  lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Color code yearly change
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

    'Create new summary table for min max

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    'pull in stock ticker to summary table when it finds min max percent change and max volume
        For i = 2 To lastrow_summary_table
        
            'Max percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Min percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Max volume
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
  
    
    Next ws
        
End Sub
