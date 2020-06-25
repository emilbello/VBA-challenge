Attribute VB_Name = "Module1"
Sub Homework2()
    
    ' Create a script that will loop through all the stocks for one year and output the following information.
        ' The ticker symbol.
        ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The total stock volume of the stock.
    ' You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
                
    
    'loop thtrough Worksheets
    
    For Each ws In Worksheets  'remember to add the ws to each cell reference
        
        ' set the table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
    
        'Dim year_change As Double
        'Dim pct_change As Double
        
        ' define variable holding the ticker
        Dim ticker As String
        
        ' declare the variable holdingthe open/close price
        Dim openprice As Double
        Dim closeprice As Double
        Dim pct_change As Double
        ' declare the variable holding the total volume and set initial value for the counter
        Dim total_volume As LongLong
        total_volume = 0
        
        ' counter for number of rows
        Dim icounter As Integer
        icounter = 0
        
        ' define new tablerow starting point
        Dim new_table_row As Long
        new_table_row = 2
        
        ' find the last row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        ' Loop through tickers
        For I = 2 To lastRow
            
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                
                icounter = icounter + 1
                
                       
                ' set ticker name
                ticker = ws.Cells(I, 1).Value
                ' add total to each ticker
                total_volume = total_volume + ws.Cells(I, 7).Value
                
                ' set open and close values
                closeprice = ws.Cells(I, 6).Value
                openprice = ws.Cells(I - icounter + 1, 3).Value
                y_change = closeprice - openprice
                
                If openprice = 0 Then
                    pct_change = 0
                Else
                    pct_change = y_change / openprice
                End If
                          
                ' write backon the new table ticker and total volume
                ws.Range("I" & new_table_row).Value = ticker
                ws.Range("L" & new_table_row).Value = total_volume
                
                
                ' write back the yearly change & percent change
                ws.Range("J" & new_table_row).Value = y_change
                ws.Range("K" & new_table_row).Value = pct_change
                
                ' incremen the new table row
                new_table_row = new_table_row + 1
                
                ' reset totals
                total_volume = 0
                icounter = 0
            
            ' When the following cell is the same ticker
            Else
                
                ' Add up total vol. Add up the counter for the number of rows
                total_volume = total_volume + ws.Cells(I, 7).Value
                icounter = icounter + 1
            
            End If
            
            
        Next I
        
        ' Format Cells autofit and percentageo pct_change
        ws.Columns("I:Q").Columns.AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
        
        ' conditional formating the y_change column
        ' declare varibles
        redcolor = 3
        greencolor = 4
        
        ' find summary table last row
        lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' iterate within y_change in the new table to format cells
        For y = 2 To lastRow2
            If ws.Cells(y, 10).Value > 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = greencolor
            ElseIf ws.Cells(y, 10).Value < 0 Then
                ws.Cells(y, 10).Interior.ColorIndex = redcolor
        
            End If
        
        Next y
        
    
    Next ws
    
End Sub




