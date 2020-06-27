Attribute VB_Name = "Module2"
Sub Challenge()

     ' CHALLENGE
        ' Your solution will also be able to return the stock with the "Greatest % increase",
        ' "Greatest % decrease" and "Greatest total volume".
       
    ' iterate through the worksheets
    For Each ws In Worksheets
       
        ' set the new table headers and field desciption
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        
        ' declare variabless
        Dim increase_pct_ticker As String
        Dim increase_pct_value As Double
        Dim decrease_pct_ticker As String
        Dim decrease_pct_value As Double
        Dim increase_total_ticker As String
        Dim increase_total_value As LongLong
        
        ' find the previous activity table last row
        lastrowTable2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' iterate to find the greatest % increase, greatest % decrease, and greatest total vol
        For n = 2 To lastrowTable2
            If ws.Cells(n, 11).Value = Application.WorksheetFunction.Max(ws.Range("K:K")) Then
                increase_pct_ticker = ws.Cells(n, 9).Value
                increase_pct_value = ws.Cells(n, 11).Value
                ' return values to cells
                ws.Range("P2").Value = increase_pct_ticker
                ws.Range("Q2").Value = increase_pct_value
                
             End If
            
            If ws.Cells(n, 11).Value = Application.WorksheetFunction.Min(ws.Range("K:K")) Then
                decrease_pct_ticker = ws.Cells(n, 9).Value
                decrease_pct_value = ws.Cells(n, 11).Value
                ' return values to cells
                ws.Range("P3").Value = decrease_pct_ticker
                ws.Range("Q3").Value = decrease_pct_value
                
            End If
            
            If ws.Cells(n, 12).Value = Application.WorksheetFunction.Max(ws.Range("L:L")) Then
                increase_total_ticker = ws.Cells(n, 9).Value
                increase_total_value = ws.Cells(n, 12).Value
                ' return values to cells
                ws.Range("P4").Value = increase_total_ticker
                ws.Range("Q4").Value = increase_total_value
                
            End If
            
                      
        Next n
            
        ' format table
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "#,###,##0"
        ws.Columns("O:Q").Columns.AutoFit
        
    Next ws
        
End Sub

