Sub HomeWork2()
    
For Each ws In Worksheets

       
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Dim ncnt As Integer
    Dim opn As Double
    Dim cls As Double
    Dim countItem As Double
    
    
    Dim Yearly_change As Double
    ncnt = 2
    opn = 0
    cls = 0
    countItem = 0
    

        
    Total_Rows = ws.Cells(Rows.Count, "A").End(xlUp).Row
   
    
    For i = 1 To Total_Rows
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(ncnt, 9).Value = ws.Cells(i + 1, 1).Value
            opn = ws.Cells(i + 1, 3).Value
            countItem = WorksheetFunction.CountIf(ws.Range("A:A"), ws.Cells(ncnt, 9).Value)
            cls = ws.Cells(i + countItem, 6).Value
            ws.Cells(ncnt, 10).Value = cls - opn
            ws.Cells(ncnt, 10).NumberFormat = "0.000000000"
            If opn = 0 Then
                ws.Cells(ncnt, 11).Value = 0
            Else
                ws.Cells(ncnt, 11).Value = (cls - opn) / opn
            End If
            
            ncnt = ncnt + 1
        End If
            
    Next i
    
    tot_rpt = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    
    For cnt = 2 To tot_rpt
      
      If ws.Cells(cnt, 10).Value >= 0 Then
        ws.Cells(cnt, 10).Interior.ColorIndex = 4
      Else
        ws.Cells(cnt, 10).Interior.ColorIndex = 3
      End If
      
    Dim arange As Range
    Dim crange As Range
    Dim grange As Range
    
    Set arange = ws.Range("A:A")
    Set crange = ws.Range("C:C")
    Set grange = ws.Range("G:G")
    
      
      ws.Cells(cnt, 12).Value = WorksheetFunction.SumIfs(grange, arange, ws.Cells(cnt, 9).Value)
         
    Next cnt
    
    
    ws.Columns("K").NumberFormat = "0.00%"
    
    
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
   
    
    
     ws.Cells(2, 18).Value = WorksheetFunction.Max(ws.Range("K:K"))
     ws.Cells(3, 18).Value = WorksheetFunction.Min(ws.Range("K:K"))
     ws.Cells(4, 18).Value = WorksheetFunction.Max(ws.Range("L:L"))
     
     For cnt3 = 1 To tot_rpt
        If ws.Cells(cnt3, 11).Value = ws.Cells(2, 18).Value Then
            ws.Cells(2, 17).Value = ws.Cells(cnt3, 9).Value
        ElseIf ws.Cells(cnt3, 11).Value = ws.Cells(3, 18).Value Then
            ws.Cells(3, 17).Value = ws.Cells(cnt3, 9).Value
        ElseIf ws.Cells(cnt3, 12).Value = ws.Cells(4, 18).Value Then
            ws.Cells(4, 17).Value = ws.Cells(cnt3, 9).Value
        End If
     Next cnt3
     
     ws.Cells(2, 18).NumberFormat = "0.00%"
     ws.Cells(3, 18).NumberFormat = "0.00%"
     ws.Cells(4, 18).NumberFormat = "0.00"
         
    
Next ws
   
        
End Sub