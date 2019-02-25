Sub Stock_Analysis()
    'count tracks change in ticker or year for a ticker and for output to next row
    'tic_year tracks the year of ticker for volume
    Dim i, count, tic_year As Integer
    'total_volume keep track of volume for ticker in a year
    Dim total_volume As Double
    'last_row is used for looping till the end of sheet
    Dim last_row As Long
    'Varible to track each ticker
    Dim ticker_symbol As String
    
For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        MsgBox (WS.Name)
        'Finding the last row for each sheet
         last_row = Cells(Rows.count, 1).End(xlUp).Row
   
        'setting the count to 0 for first row
        count = 0
        'setting the header for output
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Year"
        Cells(1, 11).Value = "Total Volume"
        
        
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
                                
            'writing the output to cells when ticker changes and intializes the values for next ticker on sheet
            ElseIf count <> 0 And ticker_symbol <> Cells(i, 1).Value Then
                'Write the ticker symbol and total volume to cell
                Cells(count + 1, 9).Value = ticker_symbol
                Cells(count + 1, 10).Value = tic_year
                Cells(count + 1, 11).Value = total_volume
                'restting the total volume and ticker_syblo,start year for next ticker
                total_volume = 0
                ticker_symbol = Cells(i, 1).Value
                count = count + 1
                tic_year = Left(Cells(i, 2).Value, 4)
            End If
            
            'If same ticker then calculate the total volume
            If ticker_symbol = Cells(i, 1).Value Then
                'If same year update the total volme
                If tic_year = Left(Cells(i, 2).Value, 4) Then
                    total_volume = total_volume + Cells(i, 7).Value
                Else
                    'Write the ticker symbol and total volume to cell when year changes
                    Cells(count + 1, 9).Value = ticker_symbol
                    Cells(count + 1, 10).Value = tic_year
                    Cells(count + 1, 11).Value = total_volume
                     'restting the total volume and ticker_syblo,start year for next ticker
                    total_volume = Cells(i, 7).Value
                    count = count + 1
                    tic_year = Left(Cells(i, 2).Value, 4)
                    
                End If
    
            End If
            'Write thelast ticker symbol and total volume to cell
            If i = last_row Then
                Cells(count + 1, 9).Value = ticker_symbol
                Cells(count + 1, 10).Value = tic_year
                Cells(count + 1, 11).Value = total_volume
            End If
            
        Next i
Next WS
End Sub
