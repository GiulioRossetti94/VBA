Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub calcu()
Application.Calculation = xlCalculationAutomatic
End Sub
Sub Open_Outlook()

Dim Ret As Long
On Error GoTo aa
Ret = ShellExecute(Application.hwnd, vbNullString, "Outlook", vbNullString, "C:\", SW_SHOWNORMAL)
If Ret < 3 Then

MsgBox "Outlook is not found.", vbCritical, "SN's Customised Solutions"
End If
aa:
End Sub

Public Function GetColor(r As Range) As Integer
GetColor = r.Interior.ColorIndex
End Function

Sub format_general()

'Sheets("Foglio3").UsedRange.NumberFormat = "General"
Sheets("index").Range("A:A").NumberFormat = "yyymmdd"


End Sub

Sub date_border()

For i = 540 To 646
    If Month(cells(i, 9)) <> Month(cells(i - 1, 9)) Then
        With Range(cells(i, 9), cells(i, 16)).Borders(xlEdgeTop)
            .LineStyle = xlcontinuos
            .Weight = xlThin
        End With
    End If
'    Debug.Print Month(Cells(i, 9))
Next

End Sub


Sub put_NAN()


For i = 11 To 60

    string_f = Replace(cells(4, i).Formula, "=+", "")
    cells(4, i).Formula = "=IF(" & string_f & "=""n.d."",""nan""," & string_f & ")"
'    c = "=IF(" & string_f & "=""n.d."",""nan""," & string_f & ")"
    Debug.Print C
        
Next i

End Sub

Sub put_year()

For i = 27 To 53

    string_f = Replace(cells(3, i).Formula, Right(cells(3, i).Formula, 4), "")
    cells(3, i).Formula = "=" & Chr(34) & string_f & Chr(34) & "&" & cells(1, i).Address


Next i


End Sub

Function count_if_colourText(colourCell As Range, area As Range)

Dim text_col As Double
Dim sum_colour As Double

text_col = colourCell.Font.Color
sum_colour = 0
For Each cell In area

    i_col = cell.Font.Color
    
    If i_col = text_col Then
        sum_colour = sum_colour + cell.Value
    End If
Next cell

count_if_colourText = sum_colour

End Function



