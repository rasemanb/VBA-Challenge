Attribute VB_Name = "Module1"
Sub stocks()

For Each ws In Worksheets
    Dim open_amt As Double
    Dim close_amt As Double
    Dim yearly As Double
    Dim percentage As Double
    Dim r As Long: r = 2
    Dim summary As Long: summary = 2
    Dim vol As Double
   
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "yearly_change"
    ws.Cells(1, 11).Value = "percent_change"
    ws.Cells(1, 12).Value = "total_stock_volume"
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        vol = WorksheetFunction.Sum(Range(ws.Cells(summary, 7), ws.Cells(i, 7)))
        open_amt = ws.Cells(summary, 3).Value
        close_amt = ws.Cells(i, 6).Value
        yearly = close_amt - open_amt
        ticker = ws.Cells(i, 1).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ws.Cells(r, 9).Value = ticker
            ws.Cells(r, 10).Value = yearly
            
            If ws.Cells(summary, 3).Value <> 0 Then
            
                percentage = yearly / open_amt
                ws.Cells(r, 11).Value = Format(percentage, "Percent")
            
            Else
            
                ws.Cells(r, 11).Value = Format(percentage, "Percent")

            End If
        
            If ws.Cells(r, 10).Value < 0 Then
        
                ws.Cells(r, 10).Interior.ColorIndex = 3
                
            Else
        
                ws.Cells(r, 10).Interior.ColorIndex = 4
                
            End If
        
        ws.Cells(r, 12).Value = vol
        r = r + 1
        summary = i + 1
        
        End If
        
    Next i
Next ws


End Sub
