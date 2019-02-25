Sub Stock_Analysis()
    'count tracks change in ticker or year for a ticker and for output to next row
    'tic_year tracks the year of ticker for volume
    Dim i, count, tic_year As Integer
    'total_volume keep track of volume for ticker in a year
    Dim total_volume, open_price, close_price, yearly_change, percent_change As Double
    Dim max_percent_increase, max_percent_decrease, max_vol As Double
    'last_row is used for looping till the end of sheet
    'last_row_output tracks the last row of calculated data
    'row_num_max, row_num_min, row_max_vol tracks row # of max % change,min % change and max volumme
    Dim last_row, last_row_output, row_num_max, row_num_min, row_max_vol As Long
    'Varible to track each ticker
    Dim ticker_symbol As String
    
For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        'MsgBox (WS.Name)
        'Finding the last row for each sheet
         last_row = Cells(Rows.count, 1).End(xlUp).Row
   
        'setting the count to 0 for first row
        count = 0
        'setting the header for output
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
 
       'starting i from 2 to ignore the header
        'Calculating the total ticker volume by year
         For i = 2 To last_row
            'setting the ticker symbol with the stock symbol
            'Count = 0 intializes the values for first ticker on sheet
            If count = 0 Then
                ticker_symbol = Cells(i, 1).Value
                count = count + 1
                total_volume = 0
                tic_year = Left(Cells(i, 2).Value, 4)
                'Getting the opening price and closing price for stock for that year
                open_price = Cells(i, 3).Value
                close_price = Cells(i, 6).Value
                                
            'writing the output to cells when ticker changes and intializes the values for next ticker on sheet
            ElseIf count <> 0 And ticker_symbol <> Cells(i, 1).Value Then
                close_price = Cells(i - 1, 6).Value
                yearly_change = close_price - open_price
                'Write the data for ticker before resetting variable for next one
                Cells(count + 1, 9).Value = ticker_symbol
                Cells(count + 1, 10).Value = yearly_change
                Cells(count + 1, 10).NumberFormat = "0.00000000"
                'filling the cell with background color green for positive change and Red for negative change
                If yearly_change > 0 Then
                    Cells(count + 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(count + 1, 10).Interior.ColorIndex = 3
                End If
                '
                If open_price > 0 Then
                    percent_change = (yearly_change / open_price)
                Else
                    percent_change = 0
                End If
                
                Cells(count + 1, 11).Value = percent_change
                Cells(count + 1, 11).NumberFormat = "0.00%"
                Cells(count + 1, 12).Value = total_volume
                'restting the total volume and ticker_syblo,start year for next ticker
                total_volume = 0
                ticker_symbol = Cells(i, 1).Value
                count = count + 1
                tic_year = Left(Cells(i, 2).Value, 4)
                open_price = Cells(i, 3).Value
                close_price = Cells(i, 6).Value
            End If
            
            'If same tickera and same year then calculate the total volume
            If ticker_symbol = Cells(i, 1).Value Then
                'If same year update the total volme
                If tic_year = Left(Cells(i, 2).Value, 4) Then
                    total_volume = total_volume + Cells(i, 7).Value
                Else
                    ' write the data and reset the variables for next year
                    close_price = Cells(i - 1, 6).Value
                    yearly_change = close_price - open_price
                    'Write the ticker symbol and total volume to cell
                    Cells(count + 1, 9).Value = ticker_symbol
                    Cells(count + 1, 10).Value = yearly_change
                    Cells(count + 1, 10).NumberFormat = "0.00000000"
                    'filling the cell with background color green for positive change and Red for negative change
                    If yearly_change > 0 Then
                         Cells(count + 1, 10).Interior.ColorIndex = 4
                    Else
                        Cells(count + 1, 10).Interior.ColorIndex = 3
                    End If
                    If open_price > 0 Then
                        percent_change = (yearly_change / open_price)
                    Else
                        percent_change = 0
                    End If
                    Cells(count + 1, 11).Value = percent_change
                    Cells(count + 1, 11).NumberFormat = "0.00%"
                    Cells(count + 1, 12).Value = total_volume
                     'restting the total volume and ticker_syblo,start year for next ticker
                    total_volume = Cells(i, 7).Value
                    count = count + 1
                    tic_year = Left(Cells(i, 2).Value, 4)
                    open_price = Cells(i, 3).Value
                    close_price = Cells(i, 6).Value
                    
                End If
    
            End If
            'Write thelast ticker symbol and total volume to cell
            If i = last_row Then
                close_price = Cells(i - 1, 6).Value
                yearly_change = close_price - open_price
                'Write the ticker symbol and total volume to cell
                Cells(count + 1, 9).Value = ticker_symbol
                Cells(count + 1, 10).Value = yearly_change
                Cells(count + 1, 10).NumberFormat = "0.00000000"
                'filling the cell with background color green for positive change and Red for negative change
                If yearly_change > 0 Then
                     Cells(count + 1, 10).Interior.ColorIndex = 4
                Else
                    Cells(count + 1, 10).Interior.ColorIndex = 3
                End If
                If open_price > 0 Then
                    percent_change = (yearly_change / open_price)
                Else
                    percent_change = 0
                End If
                Cells(count + 1, 11).Value = percent_change
                Cells(count + 1, 11).NumberFormat = "0.00%"
                Cells(count + 1, 12).Value = total_volume
            End If
            
        Next i
        
        'Calculating the last row for the output cells
        last_row_output = Cells(Rows.count, 9).End(xlUp).Row
        
        'Hard part Calculating the max % increase,decrease and volume
        
        For i = 2 To last_row_output
            If i = 2 Then
                max_percent_increase = Cells(i, 11).Value
                max_percent_decrease = Cells(i, 11).Value
                max_vol = Cells(i, 12).Value
                row_num_max = 2
                row_num_min = 2
                row_max_vol = 2
            Else
                If (max_percent_increase < Cells(i, 11).Value) Then
                    max_percent_increase = Cells(i, 11).Value
                    row_num_max = i
                End If
                If (max_percent_decrease > Cells(i, 11).Value) Then
                    max_percent_decrease = Cells(i, 11).Value
                    row_num_min = i
                End If
                If (max_vol < Cells(i, 12).Value) Then
                    max_vol = Cells(i, 12).Value
                    row_max_vol = i
                End If
                
                
            End If
            
        Next i
        
        'Writing the Header to cell
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Writing the max % increase to sheet
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(2, 16).Value = Cells(row_num_max, 9).Value
        Cells(2, 17).Value = Cells(row_num_max, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
        
         'Writing the max % decrease to sheet
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(3, 16).Value = Cells(row_num_min, 9).Value
        Cells(3, 17).Value = Cells(row_num_min, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
        
        'Writing the max volume to sheet
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(4, 16).Value = Cells(row_max_vol, 9).Value
        Cells(4, 17).Value = Cells(row_max_vol, 12).Value
       
Next WS
End Sub


