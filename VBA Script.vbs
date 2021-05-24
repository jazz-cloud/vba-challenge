Sub stockmaket():
    
    ' Set countworksheet equal to the number of worksheets in the active workbook.
    Dim countworksheet As Integer
    Dim h As Integer
    
    countworksheet = ActiveWorkbook.Worksheets.Count
    
    For h = 1 To countworksheet
        
        Worksheets(h).Activate
        
        Dim ticker_symbol As String
        Dim firstdayopenvalue As Double
        Dim lastdayclosevalue As Double
        Dim vol_total As Double
        Dim summary_table_row As Integer
        Dim summary_table_row_2 As Integer
        
        Dim ticker_with_greatest_increase As String
        Dim greatest_increase As Double
        Dim ticker_with_greatest_decrease As String
        Dim greatest_decrease As Double
        Dim ticker_with_greatest_total_volume As String
        Dim greatest_total_volume As Double
        
        ticker_symbol = "0"
        firstdayopenvalue = 0
        lastdayclosevalue = 0
        vol_total = 0
        summary_table_row = 1
        summary_table_row_2 = 1
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ticker_with_greatest_increase = "0"
        greatest_increase = 0
        ticker_with_greatest_decrease = "0"
        greatest_decrease = 0
        ticker_with_greatest_total_volume = "0"
        greatest_total_volume = 0
        
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Loop through each row
        For i = 2 To lastrow

            ' Check if still within the current ticker symbol, if it is not...
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Or i = lastrow Then
                     
                If i = lastrow Then
                
                    ' ==================== Process the last ticker for the last day data
                    ' Obtain last day close value
                    lastdayclosevalue = Cells(i, 6).Value
                
                ElseIf ticker_symbol <> "0" Then

                    ' ==================== Process the current ticker for the last day data
                    ' Obtain last day close value
                    lastdayclosevalue = Cells(i - 1, 6).Value

                End If
                    
                ' Calculate yearly change
                yearly_change = lastdayclosevalue - firstdayopenvalue

                If firstdayopenvalue <> 0 Then

                    'Calcuate percent change
                    percent_change = (lastdayclosevalue - firstdayopenvalue) / firstdayopenvalue

                Else
                    percent_change = 0

                End If
                
                ' Print output on the summary table
                If ticker_symbol <> "0" Then
                    
                    ' Print the ticker symbol to the summary table
                    Range("I" & summary_table_row).Value = ticker_symbol
                    
                    ' Print the yearly change to the summary table
                    Range("J" & summary_table_row).Value = yearly_change

                    'Set conditional formatting that will highlight positive change in green and negative change in red.
                    If yearly_change < 0 Then

                    ' Set the cell colors to red
                    Range("J" & summary_table_row).Interior.ColorIndex = 3

                    Else

                    ' Set the cell colors to green
                    Range("J" & summary_table_row).Interior.ColorIndex = 4

                    End If
                    
                    ' Print the percent change to the summary table
                    'percent_change.NumberFormat = "0.00%"
                    Range("K" & summary_table_row).Value = percent_change
                    Range("K" & summary_table_row).NumberFormat = "0.00%"
    
                    ' Print the volume amount to the summary table
                    Range("L" & summary_table_row).Value = vol_total
                    
                    ' Check if the current ticker has the greatest increase or greatest decrease or greatest total volume
                    If ticker_with_greatest_increase = "0" Then
                        ticker_with_greatest_increase = ticker_symbol
                        greatest_increase = percent_change
                    ElseIf percent_change > greatest_increase Then
                        ticker_with_greatest_increase = ticker_symbol
                        greatest_increase = percent_change
                    End If
                    
                    If ticker_with_greatest_decrease = "0" Then
                        ticker_with_greatest_decrease = ticker_symbol
                        greatest_decrease = percent_change
                    ElseIf percent_change < greatest_decrease Then
                        ticker_with_greatest_decrease = ticker_symbol
                        greatest_decrease = percent_change
                    End If
                    
                    If ticker_with_greatest_total_volume = "0" Then
                        ticker_with_greatest_total_volume = ticker_symbol
                        greatest_total_volume = vol_total
                    ElseIf vol_total > greatest_total_volume Then
                        ticker_with_greatest_total_volume = ticker_symbol
                        greatest_total_volume = vol_total
                    End If
                
                End If

                ' ==================== Processing the new ticker for the first day data
                ' Set the ticker symbol
                ticker_symbol = Cells(i, 1).Value

                ' Obtain first day open value and but not last day close value
                firstdayopenvalue = Cells(i, 3).Value
                lastdayclosevalue = 0

                ' Reset the volume total
                vol_total = 0

                ' Add to the volume total
                vol_total = vol_total + Cells(i, 7).Value
                
                ' Add one to the summary table row
                summary_table_row = summary_table_row + 1

                ' If the cell immediately following a row is the same ticker...
            Else
                ' Add to the volume total
                vol_total = vol_total + Cells(i, 7).Value

            End If

        Next i
        
     ' Print the greatest increase, greatest decrease and greatest total volume
    Cells(2, 16).Value = ticker_with_greatest_increase
    Cells(2, 17).Value = greatest_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = ticker_with_greatest_decrease
    Cells(3, 17).Value = greatest_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 16).Value = ticker_with_greatest_total_volume
    Cells(4, 17).Value = greatest_total_volume
   
    Next h
    
End Sub
