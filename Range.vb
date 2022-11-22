Sub EnterRow()


    Dim c As Excel.Range
     
    For Each c In Selection
        c.Value = c.Row
    Next c
    
End Sub
