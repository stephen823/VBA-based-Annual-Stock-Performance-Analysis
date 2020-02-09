Sub stock_analysis():
For Each Page In Worksheets
    Page.Range("I1") = "Ticker"
    Page.Range("J1") = "Yealy Change"
    Page.Range("K1") = "Percent Change"
    Page.Range("L1") = "Total Stock Volume"
    Page.Range("O2") = "Greatest % Increase"
    Page.Range("O3") = "Greatest % Decrease"
    Page.Range("O4") = "Greatest Total Volume"
    Page.Range("P1") = "Ticker"
    Page.Range("Q1") = "Value"
    LastRow = Page.Cells(Rows.Count, 1).End(xlUp).Row
    Y = 2
    J = 2
    For I = 2 To LastRow
       If Page.Cells(LastRow, 1).Value <> Page.Cells(Y - 1, 9).Value Then
          If I <= LastRow And Page.Cells(I, 1).Value <> Page.Cells(I + 1, 1).Value Then
             Page.Cells(Y, 9).Value = Page.Cells(I, 1).Value
             Page.Cells(Y, 10).Value = Page.Cells(I, 6).Value - Page.Cells(J, 3)
             Page.Cells(Y, 12).Value = Application.Sum(Range(Cells(J, 7), Cells(I, 7)))
             If Page.Cells(J, 3).Value > 0 Then
                Page.Cells(Y, 11).Value = (Page.Cells(Y, 10).Value / Page.Cells(J, 3).Value) * 100
             Else
                Page.Cells(Y, 11).Value = "NA"
             End If
             Page.Cells(I, 10).Value = Price - Page.Cells(I, 3).Value
             If Page.Cells(Y, 10).Value > 0 Then
             Page.Cells(Y, 10).Interior.ColorIndex = 4
             ElseIf Page.Cells(Y, 10).Value < 0 Then
             Page.Cells(Y, 10).Interior.ColorIndex = 3
             End If
             Y = Y + 1
              J = I + 1
          End If
        End If
     Next I
     
    Column1 = Page.Range("K:K")
    Column2 = Page.Range("L:L")
    Maxincrease = Application.WorksheetFunction.Max(Column1)
    Maxdecrease = Application.WorksheetFunction.Min(Column1)
    Maxvolume = Application.WorksheetFunction.Max(Column2)
    
    Page.Range("Q2") = Maxincrease
    Page.Range("Q3") = Maxdecrease
    Page.Range("Q4") = Maxvolume
    
    max_in_row = Page.Range("K:K").Find(WorksheetFunction.Max(Page.Range("K:K"))).Row
    min_de_row = Page.Range("K:K").Find(WorksheetFunction.Min(Page.Range("K:K"))).Row
    max_vol_row = Page.Range("L:L").Find(WorksheetFunction.Max(Page.Range("L:L"))).Row
    
    Page.Range("P2") = Page.Cells(max_in_row, 1).Value
    Page.Range("P3") = Page.Cells(min_de_row, 1).Value
    Page.Range("P4") = Page.Cells(max_vol_row, 1).Value
Next Page
End Sub


