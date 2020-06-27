Attribute VB_Name = "Module1"
Sub HomeworkVBA()
    
    ' Create a script that will loop through all the stocks for one year and output the following information.
        ' The ticker symbol.
        ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The total stock volume of the stock.
    ' You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    ' SUBMISSION COMMENTS: special thanks forTA Farshad for the pseudo code and guidelines to complete the open price loop
                
    
    'loop thtrough Worksheets
    For Each ws In Worksheets
        
        ' set the new table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ' declare variable holding the ticker
        Dim ticker As String
        
        ' declare the variable holdingthe open/close price and pct change
        Dim openprice As Double
        Dim closeprice As Double
        Dim pct_change As Double
        ' declare the variable holding the total volume and set initial value for the counter
        Dim total_volume As LongLong
        total_volume = 0
        
        ' set row openprice for the ticker
        Start = 2
        
        ' declare/define new tablerow starting point
        Dim new_table_row As Long
        new_table_row = 2
        
        ' find the last row in column A, tickers
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Loop through tickers
        For I = 2 To lastRow
            ' check if next cell is different for ticker value
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                
               ' set ticker name
                ticker = ws.Cells(I, 1).Value
                ' add total to each ticker
                total_volume = total_volume + ws.Cells(I, 7).Value
                
                ' looking for the first non-zero openprice, when total vol<>0
                If total_volume = 0 Then
                    ws.Range("J" & new_table_row).Value = 0
                    ws.Range("K" & new_table_row).Value = 0
                
                Else
                    If ws.Cells(Start, 3).Value = 0 Then
                        For x = Start To I
                            If ws.Cells(x, 3).Value <> 0 Then
                                Start = x
                                Exit For
                            End If
                        Next x
                    End If
                    ' set open and close reference values
                    closeprice = ws.Cells(I, 6).Value
                    openprice = ws.Cells(Start, 3).Value
                    y_change = closeprice - openprice
                    pct_change = y_change / openprice
                    
                    ' write back the yearly change & percent change
                    ws.Range("J" & new_table_row).Value = y_change
                    ws.Range("K" & new_table_row).Value = pct_change
                    
                    Start = I + 1
                 End If
                    ' set new start row for the openprice
                                 
                              
                ' write back on the new table ticker and total volume
                ws.Range("I" & new_table_row).Value = ticker
                ws.Range("L" & new_table_row).Value = total_volume
                
                ' increment the new table row
                new_table_row = new_table_row + 1
                ' reset totals
                total_volume = 0
                icounter = 0
            
            ' When the following cell is  different ticker
            Else
                
                ' Add up total vol. Add up the counter for the number of rows ticker
                total_volume = total_volume + ws.Cells(I, 7).Value
                icounter = icounter + 1
            
            End If
            
            
        Next I
        
                
        ' conditional formating the y_change column
        ' declare varibles
        redcolor = 3
        greencolor = 4
        
        ' find summary table last row, in columnJ
        lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' iterate to format Yearly Change cells in green and red as per the value
        For y = 2 To lastRow2
            If ws.Cells(y, 10).Value > 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = greencolor
            ElseIf ws.Cells(y, 10).Value < 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = redcolor
            End If
        Next y
        
       ' Format percentage for pct_change and thousand separator for total_volume
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "#,###,##0"
        ' autofit all colmns in the new tables
        ws.Columns("I:L").Columns.AutoFit
        
    Next ws
    
End Sub

