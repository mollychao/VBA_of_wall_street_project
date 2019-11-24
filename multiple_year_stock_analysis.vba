Sub mutil_year_stock()
    

    Dim ws As Worksheet

    'Start loop in multiple worksheet & create the table
    For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'Dim variables
    Dim ticket As String
    Dim total_volume As Double
    total_volume = 0
    Dim rowcount As Long
    rowcount = 2
    Dim yearly_open As Double
    Dim yearly_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Long

        'looping with conditions
        For i = 2 To lastrow
         
         
        'find the yearly open
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                yearly_open = ws.Cells(i, 3).Value
            End If
        
        'add total volume from each row
            total_volume = total_volume + ws.Cells(i, 7)
            
         'find the yearly close (if the ticker is changing)
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                yearly_close = ws.Cells(i, 6).Value

                'if the ticker changes, stop adding volume
                ws.Cells(rowcount, 12).Value = total_volume
                total_volume = 0
                
                'yearly change add to the summary
                yearly_change = yearly_close - yearly_open
                ws.Cells(rowcount, 10).Value = yearly_change

                'percentage change
                'To avoid zero division
                If yearly_open = 0 And yearly_close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                ElseIf yearly_open = 0 Then
                    Dim percent_change_NA As String
                    percent_change_NA = "NA"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = yearly_change / yearly_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                End If

                If yearly_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If
            
                'reset
                rowcount = rowcount + 1
                
                yearly_open = 0
                yearly_close = 0
                yearly_change = 0
                
                percent_change = 0
                
            End If
        Next i

        '--------------------------------------------------------------
        'performance chart
        
        'create table
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Values"

        'Dim variables
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_volume_ticker As String
        
        Dim increase_percentage As Double
        increase_percentage = ws.Cells(2, 11).Value

        Dim decrease_percentage As Double
        decrease_percentage = ws.Cells(2, 11).Value

        Dim volume_count As Double
        volume_count = ws.Cells(2, 12).Value


        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            'looping with condition
            For i = 2 To lastrow
        
        
                'greatest increase
                If ws.Cells(i, 11).Value > increase_percentage Then
                    increase_percentage = ws.Cells(i, 11).Value
                    greatest_increase_ticker = ws.Cells(i, 9).Value
                    ws.Cells(2, 15).Value = greatest_increase_ticker
                    ws.Cells(2, 16).Value = increase_percentage
                    ws.Cells(2, 16).NumberFormat = "0.00%"
                End If

                'greatest decrease
                If ws.Cells(i, 11).Value < decrease_percentage Then
                    decrease_percentage = ws.Cells(i, 11).Value
                    greatest_decrease_ticker = ws.Cells(i, 9).Value
                    ws.Cells(3, 15).Value = greatest_decrease_ticker
                    ws.Cells(3, 16).Value = decrease_percentage
                    ws.Cells(3, 16).NumberFormat = "0.00%"
                End If

                'greatest volume
                If ws.Cells(i, 12).Value > volume_count Then
                    volume_count = ws.Cells(i, 12).Value
                    greatest_volume_ticker = ws.Cells(i, 9).Value
                    ws.Cells(4, 15).Value = greatest_volume_ticker
                    ws.Cells(4, 16).Value = volume_count
                End If

        Next i

    Next ws

End Sub

