Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Cells.count > 1 Then Exit Sub
    If Not Intersect(Target, Range("M28", "AZ67")) Is Nothing Then
        Target.Interior.Color = RGB(WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255))
    End If
End Sub

'##################################v2
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Range("B2", "K12")) Is Nothing Then
    n = WorksheetFunction.RandBetween(0, 255)
'        Target.Interior.Color = RGB(WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255), WorksheetFunction.RandBetween(0, 255))
        If Target.Column Mod 2 = 0 Then
            Target.Interior.Color = RGB(0, n, 0)
        Else
            Target.Interior.Color = RGB(0, 0, n)
        End If
        
        Target.Value = n
    
    End If
End Sub
