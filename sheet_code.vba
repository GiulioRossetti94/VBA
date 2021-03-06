
'======================================================================================================================
Private Sub CheckBox4_Change()
Dim x As Long
If Me.CheckBox4.Value Then
For x = 1 To 3
    Me.OLEObjects("CheckBox" & x).Object.Value = True
Next x
End If

End Sub

'======================================================================================================================
'COMMAND BUTTON NAME -> CLEAR

Private Sub CommandButton1_Click()
Range("M28", "AZ67").ClearFormats
Range("M28", "AZ67").BorderAround 1, xlThick
Range("M28", "AZ67").Font.Color = vbWhite

Range("BE28", "CR67").ClearContents
Range("BE28", "CR67").BorderAround 1, xlThick
End Sub

Private Sub CommandButton2_Click()
Dim cells As Range
'random_n = Rnd()
'
'If random_n > 0.5 Then
'    clr = 255
'Else:
'    clr = 0
'End If

For Each cells In Range("M28", "AZ67")
        If Me.CheckBox2.Value Then
            red = WorksheetFunction.RandBetween(0, 255)
        Else
            red = 0
        End If
        
        If Me.CheckBox1.Value Then
            green = WorksheetFunction.RandBetween(0, 255)
        Else
            green = 0
        End If
        
        If Me.CheckBox3.Value Then
            blue = WorksheetFunction.RandBetween(0, 255)
        Else
            blue = 0
        End If
        
        cells.Interior.Color = rgb(red, green, blue)
        
Next cells
End Sub

'======================================================================================================================
COMMAND BUTTON NAME -> PIXEL PIC

Private Sub CommandButton3_Click()
Range("BE28", "CR67").ClearContents
Range("BE28", "CR67").BorderAround 1, xlThick
Call importArray
Call delNamedRange
If Me.CheckBox8.Value Then
    Call HEXtoCOL
Else
    Call newHEXtoCOL
End If
End Sub

'======================================================================================================================
COMMAND BUTTON NAME -> SORT COLORS
Private Sub CommandButton4_Click()
Call ColorOrdering
End Sub

'======================================================================================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Range("M28", "AZ67")) Is Nothing Then
        If Me.CheckBox2.Value Then
            red = WorksheetFunction.RandBetween(0, 255)
        Else
            red = 0
        End If
        
        If Me.CheckBox1.Value Then
            green = WorksheetFunction.RandBetween(0, 255)
        Else
            green = 0
        End If
        
        If Me.CheckBox3.Value Then
            blue = WorksheetFunction.RandBetween(0, 255)
        Else
            blue = 0
        End If
        
        Target.Interior.Color = rgb(red, green, blue)
    End If
End Sub


'======================================================================================================================
