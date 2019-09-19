Sub blo_loop()

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
'newPPT.WindowState = 2

With newPPT.ActivePresentation
    .PageSetup.FirstSlideNumber = 2
End With



newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
newPPT.ActivePresentation.ApplyTemplate "C:\Users\bloomberg03\AppData\Roaming\Microsoft\Templates\FERI CTG.potx"

'list name of layout
'With newPPT.ActivePresentation
' For Each desName In .Designs
'
'            MsgBox "The design name is " & .Designs.Item(desName.Index).Name
'
'        Next
'End With

''=======================================================================================================================
''COVER
''=======================================================================================================================
Set pptLayout = newPPT.ActivePresentation.Designs(2).SlideMaster.CustomLayouts(1)
Set activeslide = newPPT.ActivePresentation.Slides.AddSlide(1, pptLayout)

activeslide.Shapes.AddShape Type:=msoShapeRectangle, _
    Left:=17, Top:=236, Width:=672, Height:=26.6
    
With activeslide.Shapes(1)
    .Fill.ForeColor.rgb = rgb(255, 255, 255)
    .Line.Visible = msoFalse
End With

With activeslide.Shapes(1).TextFrame.TextRange
    .Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Text = "Portfolio Update " & Format(Now(), "dd/mm/yyyy")
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))

Dim strFolder As String
Dim strFileName As String
'Dim objPic As Picture
'Dim rngCell As Range
'
strFolder = "C:\Users\bloomberg03\Desktop\BBL_pic"
If Right(strFolder, 1) <> "\" Then
    strFolder = strFolder & "\"
End If

strFileName = Dir(strFolder & "*.jpg", vbNormal)

For i = 1 To 10
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
    activeslide.Shapes(1).Left = 17
    activeslide.Shapes(1).Top = 24
    With activeslide.Shapes(1).TextFrame.TextRange
        If i = 1 Then
            .Text = "Performance"
            ElseIf i = 2 Then .Text = "Caratteristiche portafoglio"
            ElseIf i = 3 Then .Text = "Caratteristiche portafoglio"
            ElseIf i = 4 Then .Text = "Flussi di cassa"
            ElseIf i = 5 Then .Text = "Tassi chiave"
            ElseIf i = 6 Then .Text = "Volatilità"
            ElseIf i = 7 Then .Text = "Comparazione VaR (P&L)"
            ElseIf i = 8 Then .Text = "Comparazione VaR (rend%)"
            ElseIf i = 9 Then .Text = "Peggiori scenari Fixed Income"
            ElseIf i = 10 Then .Text = "Peggiori scenari Equity"
        End If
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
    
    strFileName = Dir(strFolder & i & "*.jpg", vbNormal)
    Set bl_im = activeslide.Shapes.AddPicture(strFolder & strFileName, msoFalse, msoTrue, 36.85, 72, 552, 397)
    strFileName = Dir(strFolder & strFileName)
    activeslide.Shapes(2).Delete
    Application.CutCopyMode = False
    
Next
''Loop



End Sub