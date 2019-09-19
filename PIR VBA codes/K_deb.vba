Sub paste_for_charts()
 lr = Worksheets("Rnd Daily").Range("BI3").End(xlDown).Row
 increment = 3
For i = 1 To 10
    With ThisWorkbook.Worksheets("Rnd Daily")
    
        cells(2, 76 + increment) = cells(1, 60 + i)
        .Range(cells(3, 61), cells(lr, 63)).Copy
        .Range(cells(3, 75 + increment), cells(lr, 77 + increment)).PasteSpecial Paste:=xlPasteFormulas
    
    End With

    increment = increment + 3

Next

'With ThisWorkbook.Worksheets("Rnd Daily")
'
'    .Range(Cells(3, 61), Cells(lr, 63)).Copy
'    .Range(Cells(3, 75), Cells(lr, 77)).PasteSpecial Paste:=xlPasteValues
'
'End With
End Sub

Sub charts_2()
Dim newPPT As PowerPoint.Application
Dim Aslide As PowerPoint.Slide
Dim i As Integer
Dim r1 As Range
Dim r2 As Range
Dim r3 As Range
Dim r4 As Range
Dim rweek As Range
Dim FileName As String
FileName = "C:\Users\bloomberg03\Desktop\Daily Market Chart\Daily Recap_" & Format(Now, "ddmmyy") & ".pdf"

Sheets("Rnd Daily").Activate
Set r1 = Range("C3:J32")
Set r2 = Range("C36:Q67")
Set r3 = Range("C70:P79")
Set r4 = Range("E84:N134")


On Error Resume Next
Set newPPT = GetObject(, "PowerPoint.Application")

On Error GoTo 0

If newPPT Is Nothing Then
    Set newPPT = New PowerPoint.Application
    End If

If newPPT.Presentations.Count = 0 Then
    newPPT.Presentations.Add (msoCTrue)
    End If
    
Set WDReport = newPPT.Presentations.Open("C:\Users\bloomberg03\Desktop\Daily Market Chart\daily market_template.pptx")
    
 Application.ScreenUpdating = False
newPPT.Visible = True

Sheets("Rnd Daily").Activate

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
newPPT.ActivePresentation.ApplyTemplate "C:\Users\bloomberg03\AppData\Roaming\Microsoft\Templates\FERI CTG.potx"
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Daily Recap"
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

r4.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 211.46
newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.425
newPPT.ActiveWindow.Selection.ShapeRange.Width = 331.08
newPPT.ActiveWindow.Selection.ShapeRange.Height = 426.33
activeslide.Shapes(2).Delete



newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Market Recap"
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

r1.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
newPPT.ActiveWindow.Selection.ShapeRange.Height = 397
activeslide.Shapes(2).Delete



newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity Portfolio"
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

r2.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 2.55
newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
newPPT.ActiveWindow.Selection.ShapeRange.Width = 710.36
newPPT.ActiveWindow.Selection.ShapeRange.Height = 373.88
activeslide.Shapes(2).Delete



newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "ETF Portfolio"
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

r3.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 7
newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
newPPT.ActiveWindow.Selection.ShapeRange.Width = 632.9763
newPPT.ActiveWindow.Selection.ShapeRange.Height = 104.5984
activeslide.Shapes(2).Delete

For i = 1 To 10

    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
    
    
    
    If i < 6 Then
        activeslide.Shapes(1).Left = 17
        activeslide.Shapes(1).Top = 24
        With activeslide.Shapes(1).TextFrame.TextRange
            .Text = "TOP 5"
            .Font.Size = 20
            .Font.Color = rgb(0, 0, 139)
            .Font.Name = "Georgia"
            .Font.Bold = True
        End With
        With activeslide.Shapes(1)

            .Left = 20.97
            .Top = 15.02
        End With
    
    
    Else:
        activeslide.Shapes(1).Left = 17
        activeslide.Shapes(1).Top = 24
        With activeslide.Shapes(1).TextFrame.TextRange
            .Text = "BOTTOM 5"
            .Font.Size = 20
            .Font.Color = rgb(0, 0, 139)
            .Font.Name = "Georgia"
            .Font.Bold = True
        End With
        With activeslide.Shapes(1)
   
            .Left = 20.97
            .Top = 15.02
        End With
    End If
    
    
    
    ActiveSheet.ChartObjects(i).Activate

    ActiveChart.ChartArea.Copy
    activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 9.637795278
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 105.44
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 702.42
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 226.77
    
    activeslide.Shapes(2).Delete
Next i

WDReport.SaveAs FileName, ppSaveAsPDF
End Sub


