
'

Sub generateReportPythonTest()
Application.DisplayAlerts = False

Dim newPPT As PowerPoint.Application
Dim Aslide As PowerPoint.Slide

On Error Resume Next
Set newPPT = GetObject(, "PowerPoint.Application")


On Error GoTo 0

If newPPT Is Nothing Then
    Set newPPT = New PowerPoint.Application
    End If

If newPPT.Presentations.Count = 0 Then
    newPPT.Presentations.Add
    End If
    
 Application.ScreenUpdating = False
newPPT.Visible = True
newPPT.WindowState = 2
 With newPPT.ActivePresentation

    .PageSetup.FirstSlideNumber = 2
    End With

    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
'    newPPT.ActiveWindow.ViewType = ppViewSlideSorter
    newPPT.ActivePresentation.ApplyTemplate "C:\Users\bloomberg03\AppData\Roaming\Microsoft\Templates\FERI CTG.potx"
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)


Dim strFolderPython As String
Dim objPicPython As Picture

strFolderPython = "C:\Users\bloomberg03\Desktop\PythonScript\Pic\"
Name_to_check = "REPORT_Python_Pics"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True

    End If
Next Sheet

If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "REPORT_Python_Pics"
End If

If Right(strFolderPython, 1) <> "\" Then
    strFolderPython = strFolderPython & "\"
End If

'''------Importing pics in sheet

With Worksheets("REPORT_Python_Pics")
    .Columns(1).ColumnWidth = 110
    .Rows.RowHeight = 400
End With

Set rngCell = Sheets("REPORT_Python_Pics").Range("A1")

strFileNamePython = Dir(strFolderPython & "*.jpg")

nPicPython = 0

Do While Len(strFileNamePython) > 0
    Debug.Print strFolderPython & strFileNamePython
    Set objPicPython = Sheets("REPORT_Python_Pics").Pictures.Insert(strFolderPython & strFileNamePython)
        With objPicPython
            .Left = rngCell.Left
            .Top = rngCell.Top
            .Height = rngCell.RowHeight
            .Placement = xlMoveAndSize
        End With
    Set rngCell = rngCell.Offset(1, 0)
    strFileNamePython = Dir
    nPicPython = nPicPython + 1
Loop

For i = 1 To nPicPython
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
        newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
        Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
    
     activeslide.Shapes(1).Left = 17
     activeslide.Shapes(1).Top = 24
    With activeslide.Shapes(1).TextFrame.TextRange
        .Text = "Performance"
        .Font.Size = 20
        .Font.Color = rgb(0, 0, 139)
        .Font.Name = "Georgia"
        .Font.Bold = True
    End With
    With activeslide.Shapes(1)
        .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
        .Left = 20.97
        .Top = 15.02
    End With

    Sheets("REPORT_Python_Pics").Activate
    Set cp = Sheets("REPORT_Python_Pics").cells(i, 1)
    cp.CopyPicture
    activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
        newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
        newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
        newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
        newPPT.ActiveWindow.Selection.ShapeRange.Height = 397


activeslide.Shapes(2).Delete
Next i
Worksheets("REPORT_Python_Pics").Delete

Application.DisplayAlerts = True


End Sub

