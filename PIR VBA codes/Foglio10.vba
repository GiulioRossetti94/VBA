Private Sub CheckBox4_Change()
Dim x As Long
If Me.CheckBox4.Value Then
For x = 1 To 3
    Me.OLEObjects("CheckBox" & x).Object.Value = True
Next x
End If

End Sub

Private Sub CommandButton1_Click()
'MsgBox "Ciao ;)"
Range("M28", "AZ67").ClearFormats
Range("M28", "AZ67").BorderAround 1, xlThick
Range("M28", "AZ67").Font.Color = vbWhite

Range("BE28", "CR67").ClearContents
Range("BE28", "CR67").BorderAround 1, xlThick
End Sub

Private Sub CommandButton10_Click()
Call posterize
End Sub

Private Sub CommandButton11_Click()
Call IncreaseColor
End Sub

Private Sub CommandButton12_Click()
Call DecreaseColor
End Sub

Private Sub CommandButton13_Click()
Call flip_pic
End Sub

Private Sub CommandButton2_Click()
'MsgBox "Ciao ;)"
Dim cells As Range
'random_n = Rnd()
'
'If random_n > 0.5 Then
'    clr = 255
'Else:
'    clr = 0
'End If

'flag = WorksheetFunction.RandBetween(0, 1)
'If flag = 0 Then b_w = 255 Else b_w = 0

If Me.ToggleButton1.Value Then
    b_w = 255
Else
    b_w = 0
End If
For Each cells In Range("M28", "AZ67")
        If Me.CheckBox2.Value Then
            red = WorksheetFunction.RandBetween(0, 255)
        Else
            red = b_w
        End If
        
        If Me.CheckBox1.Value Then
            green = WorksheetFunction.RandBetween(0, 255)
        Else
            green = b_w
        End If
        
        If Me.CheckBox3.Value Then
            blue = WorksheetFunction.RandBetween(0, 255)
        Else
            blue = b_w
        End If
        
        cells.Interior.Color = rgb(red, green, blue)
        
Next cells
End Sub

Private Sub CommandButton3_Click()
'MsgBox "Ciao ;)"
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

Private Sub CommandButton4_Click()
'MsgBox "Ciao ;)"
Call ColorOrdering
End Sub

Private Sub CommandButton5_Click()
'MsgBox "Ciao ;)"
Call ColorOrderingHSL
End Sub

Private Sub CommandButton6_Click()
'MsgBox "Ciao ;)"
Call RandomColBand
End Sub

Private Sub CommandButton7_Click()
'MsgBox "Ciao ;)"
Call circleSquare
End Sub

Private Sub CommandButton8_Click()
'MsgBox "Ciao ;)"
Call ShufflePix
End Sub

Private Sub CommandButton9_Click()
'MsgBox "Ciao ;)"
Call square_blur
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Me.ToggleButton1.Value Then
    b_w = 255
Else
    b_w = 0
End If
If Target.cells.Count > 1 Then Exit Sub
    If Not Intersect(Target, Range("M28", "AZ67")) Is Nothing Then
        If Me.CheckBox2.Value Then
            red = WorksheetFunction.RandBetween(0, 255)
        Else
            red = b_w
        End If

        If Me.CheckBox1.Value Then
            green = WorksheetFunction.RandBetween(0, 255)
        Else
            green = b_w
        End If

        If Me.CheckBox3.Value Then
            blue = WorksheetFunction.RandBetween(0, 255)
        Else
            blue = b_w
        End If

        Target.Interior.Color = rgb(red, green, blue)
    End If
End Sub