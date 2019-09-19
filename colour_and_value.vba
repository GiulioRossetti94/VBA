Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Range("B2", "K12")) Is Nothing Then
    n = WorksheetFunction.RandBetween(70, 255)
'        Target.Interior.Color = RGB(WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255))
        If Target.Column Mod 2 = 0 Then
            Target.Interior.Color = RGB(n, 0, 0)
        Else
            Target.Interior.Color = RGB(0, 0, n)
        End If
        
        Target.Value = n
    
    End If
End Sub
