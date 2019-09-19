'SCRIPT for "PORTFOLIO UPDATE" POWERPOINT PRESENTATION
'


Sub NewReport()

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

activeslide.Shapes.AddShape Type:=msoShapeRectangle, _
    Left:=17, Top:=481.89, Width:=226.772, Height:=26.6
    
With activeslide.Shapes(2)
    .Fill.ForeColor.rgb = rgb(255, 255, 255)
    .Line.Visible = msoFalse
End With

With activeslide.Shapes(2).TextFrame.TextRange
    .Paragraphs.ParagraphFormat.Alignment = ppAlignLeft
    .Text = "Report created:" & Format(Now(), "dd-MM-yyyy hh:mm:ss")
    .Font.Size = 11
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With
''=======================================================================================================================
'''''Andamento mercati -SLIDE DIVISORIA
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Andamento Mercati"
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

With activeslide.Shapes(1)
    .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Left = 17
    .Top = 236
End With
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse
activeslide.Shapes(2).Delete

newPPT.ActivePresentation.Slides.Item(2).Delete

''=======================================================================================================================
'''''SLIDE RENDIMENTI MERCATO'''''''''''''''''''''''''''''''''''''''''''''''''
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

Set ws_tables = Sheets("Tables")
Set ind = ws_tables.Range("G4:N35")
ind.CopyPicture

'copy range
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 19
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 56.04
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 340
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 340

activeslide.Shapes.Range.Align msoAlignLefts, msoFalse
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Rendimenti Mercato"
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
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse


    activeslide.Shapes(2).Left = 19
    activeslide.Shapes(2).Top = 406.2047
    activeslide.Shapes(2).Height = 94.39
    activeslide.Shapes(2).Width = 663
 Application.CutCopyMode = False
 
''=======================================================================================================================
''Rendimenti Indici
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Indici Italia"
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

Sheets("Indici Italia").Range("G8") = "ref: "
Sheets("Indici Italia").Range("H8") = Format(Now(), "dd-MM-yyyy hh:mm:ss")
Sheets("Indici Italia").Range("H8").NumberFormat = "dd-MM-yyyy hh:mm:ss"

Sheets("Indici Italia").Rows("24:29").EntireRow.Hidden = True
Set rInd = Sheets("Indici Italia").Range("A5:M31")

rInd.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = -19.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 88.44
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 698
    newPPT.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoFalse
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 218
activeslide.Shapes(2).Delete

Sheets("Indici Italia").Rows("24:29").EntireRow.Hidden = False
Application.CutCopyMode = False

''=======================================================================================================================
'''''FERI GENERALE -SLIDE DIVISORIA
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "FONDO FERI PIR: Dati Generali"
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

With activeslide.Shapes(1)
    .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Left = 17
    .Top = 236
End With
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse
activeslide.Shapes(2).Delete

 
''=======================================================================================================================
''ANDAMENTO RACCOLTA
''=======================================================================================================================
 DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
   Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

Dim ch1 As Excel.ChartObject
Dim ch2 As Excel.ChartObject


Set ws_raccolta = Sheets("Raccolta")
ws_raccolta.ChartObjects(1).Activate

ActiveChart.ChartArea.Copy
activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 21
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 60
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 370
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 193

ActiveSheet.ChartObjects(2).Activate
ActiveChart.ChartArea.Copy
activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 21
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 276
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 370
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 193

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Andamento Raccolta"
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

    activeslide.Shapes(2).Left = 402.5
    activeslide.Shapes(2).Top = 65.76
    activeslide.Shapes(2).Height = 167
    activeslide.Shapes(2).Width = 294.8


With activeslide.Shapes(2).TextFrame.TextRange
    .Text = "Al " & Format(Now(), "short Date") & " le sottoscrizioni totali nette ammontano a " & Format(Range("R12").Text, "Currency") _
    & ", grazie alle convenzioni di collocamento con Banca Finint e Banca Valsabbina." _
    & vbCrLf & "Nel mese di Gennaio: " & Format(Range("R32").Text, "Currency") _
    & vbCrLf & "Nel mese di Febbraio: " & Format(Range("R33").Text, "Currency") _
    & vbCrLf & "Nel mese di Marzo: " & Format(Range("R34").Text, "Currency") _
    & vbCrLf & "Nel mese di Aprile: " & Format(Range("R35").Text, "Currency") _
    & vbCrLf & "Nel mese di Maggio: " & Format(Range("R36").Text, "Currency")
'    & vbCrLf & "Nel mese di Giugno: " & Format(Range("R25").Text, "Currency") _
'    & vbCrLf & "Nel mese di Luglio: " & Format(Range("R26").Text, "Currency") _
'    & vbCrLf & "Nel mese di Agosto: " & Format(Range("R27").Text, "Currency") _
'    & vbCrLf & "Nel mese di Settembre: " & Format(Range("R28").Text, "Currency") _
'    & vbCrLf & "Nel mese di Ottobre: " & Format(Range("R29").Text, "Currency") _
'    & vbCrLf & "Nel mese di Novembre: " & Format(Range("R30").Text, "Currency") _
'    & vbCrLf & "Nel mese di Dicembre: " & Format(Range("R31").Text, "Currency") _
'    & vbCrLf & "Raccolta 2017: " & Format(Range("o2").Text, "Currency") _
'   & vbCrLf & "Raccolta 2018: " & Format(Range("p2").Text, "Currency") _
'    & vbCrLf & "Raccolta 2019: " & Format(Range("q2").Text, "Currency")

    .Font.Size = 11
    .Font.Color = rgb(0, 0, 0)
    .Font.Name = "Georgia"
    .Font.Bold = False
End With

Set racc = Sheets("Raccolta").Range("N1:Q2")

racc.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 21
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 474.23
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 237.54
    newPPT.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoFalse
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 22.67
Application.CutCopyMode = False


''=======================================================================================================================
''Asset allocation
''=======================================================================================================================
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

Set ws_Performance = Sheets("Performance")
ws_Performance.ChartObjects(1).Activate

ActiveChart.ChartArea.Copy
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))

activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 280.7
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 86
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 400
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 154

activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Asset Allocation (1/2)"
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

activeslide.Shapes(2).Left = 17
activeslide.Shapes(2).Top = 54
With activeslide.Shapes(2).TextFrame.TextRange
    .Text = "Di seguito lasset allocation dellintero portafoglio (azionario + obbligazionario)."
    .Font.Size = 11
    .Font.Color = rgb(0, 0, 0)
    .Font.Name = "Georgia"
    .Font.Bold = False
End With

Set Nav = Sheets("Asset Allocation").Range("A3:B10")
Nav.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 31.18
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 85.03
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 198
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 100

Set esp = Sheets("Asset Allocation").Range("A12:D32")
esp.CopyPicture

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 31.18
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 250
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 381.82
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 216.28

Set tas = Sheets("Performance").Range("AG1:AK15")
tas.CopyPicture

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 382
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 250
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 180
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 154.5
Application.CutCopyMode = False

Set tas2019 = Sheets("Performance").Range("AN1:AP15")
tas2019.CopyPicture

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 529
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 250
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 180
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 154.5
Application.CutCopyMode = False


''=======================================================================================================================
''Asset allocation 2/2
''=======================================================================================================================
'
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Asset Allocation (2/2)"
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

activeslide.Shapes(2).Left = 17
activeslide.Shapes(2).Top = 54
With activeslide.Shapes(2).TextFrame.TextRange
    .Text = "Di seguito la diversificazione settoriale dellintero portafoglio."
    .Font.Size = 11
    .Font.Color = rgb(0, 0, 0)
    .Font.Name = "Georgia"
    .Font.Bold = False
End With

Set pie = Sheets("Asset Allocation").Range("F1:L14")
pie.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 17
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 88.5
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 674.64
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 183

'Set pie1 = Sheets("Asset Allocation").Range("F15:N28")
'pie1.CopyPicture
'Activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
'    newPPT.ActiveWindow.Selection.ShapeRange.Left = -65.2
'    newPPT.ActiveWindow.Selection.ShapeRange.Top = 88.5
'    newPPT.ActiveWindow.Selection.ShapeRange.Width = 674.64
'    newPPT.ActiveWindow.Selection.ShapeRange.Height = 183


Set ws_AA = Sheets("Asset Allocation")
ws_AA.ChartObjects(2).Activate

ActiveChart.ChartArea.Copy
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 20
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 270
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 315
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 191

ws_AA.ChartObjects(1).Activate

ActiveChart.ChartArea.Copy
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))

activeslide.Shapes.PasteSpecial(DataType:=ppPasteJPG).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 270
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 236
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 430
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 230
    newPPT.ActiveWindow.Selection.ShapeRange.ZOrder msoSendToBack
 Application.CutCopyMode = False
 
''=======================================================================================================================
''Portfolio 1
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Portafoglio (1/2)"
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

Sheets("Asset Allocation").Columns("A").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("C").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("L").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("N").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("W").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("Y").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("AH").EntireColumn.Hidden = True
Sheets("Asset Allocation").Columns("AJ").EntireColumn.Hidden = True

'Sheets("Asset Allocation").Activate
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
Set Equity = Sheets("Asset Allocation").Range("A37:J76")
Equity.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
'    newPPT.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoFalse
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 601.7953
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 377.5748

'lRow2 = Cells(118, 1).End(xlUp).Offset(1, 0).Row
'Set etf = Sheets("Asset Allocation").Range(Cells(118, "A"), Cells(lRow2, "I"))
Set etf = Sheets("Asset Allocation").Range("L37:U44")
etf.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 435
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 83

activeslide.Shapes(2).Delete
Application.CutCopyMode = False

''=======================================================================================================================
''Portfolio 2
''=======================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Portafoglio (2/2)"
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


'Sheets("Asset Allocation").Activate

'lRow2 = Cells(118, 1).End(xlUp).Offset(1, 0).Row
'Set fi_ = Sheets("Asset Allocation").Range("A75:I110")

Set fi_ = Sheets("Asset Allocation").Range("W37:AF73")
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
fi_.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 377.5748

Set gvt = Sheets("Asset Allocation").Range("AH37:AQ43")
gvt.CopyPicture
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 435
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 610
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 73

Sheets("Asset Allocation").Columns("A").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("C").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("L").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("N").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("W").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("Y").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("AH").EntireColumn.Hidden = False
Sheets("Asset Allocation").Columns("AJ").EntireColumn.Hidden = False

activeslide.Shapes(2).Delete
Application.CutCopyMode = False

''=======================================================================================================================
'''''ANALISI EQUITY -SLIDE DIVISORIA
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Portafoglio Azionario - Analisi"
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

With activeslide.Shapes(1)
    .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Left = 17
    .Top = 236
End With
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse
activeslide.Shapes(2).Delete

''=======================================================================================================================
''Rendimenti Equity & ETF
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Rendimenti Equity & ETF "
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

Set re = ws_tables.Range("Q3:AG39")

re.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 25.51
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 676.06
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 374.455

Set rETF = ws_tables.Range("AQ3:BD9")

rETF.CopyPicture
'DoEvents
'    Application.Wait (Now + TimeValue("0:00:001"))
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 7.93
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 415.874
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 683.71
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 85

activeslide.Shapes(2).Delete

'Call generateReportPythonTest

Application.CutCopyMode = False

''=======================================================================================================================
''EQUITY distribuzione settoriale
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity - Distribuzione Settoriale "
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

ts_weight_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\TS_weights.emf"
Set ts_w = activeslide.Shapes.AddPicture(ts_weight_path, msoFalse, msoTrue, 223, 100, 487.55, 281.76)

industry_risk_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\Industry_risk.emf"
activeslide.Shapes.AddPicture(industry_risk_path, msoFalse, msoTrue, -119, 77.1, 450.7087, 198.7087).ZOrder msoSendBackward


industry_alloc_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\Industry_allocation.emf"
activeslide.Shapes.AddPicture(industry_alloc_path, msoFalse, msoTrue, -119, 300.7, 450.7087, 198.7087).ZOrder msoSendBackward

Application.CutCopyMode = False
activeslide.Shapes(2).Delete

''=======================================================================================================================
''EQUITY Correlazione
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity - Distribuzione Settoriale "
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

correlation_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\correlation.emf"
activeslide.Shapes.AddPicture(correlation_path, msoFalse, msoTrue, 0, 52, 462.04, 325.7).ZOrder msoSendBackward

ml_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\mstructure_ret.emf"
Set ml = activeslide.Shapes.AddPicture(ml_path, msoFalse, msoTrue, 382.67, 202.39, 313.22, 283)

Application.CutCopyMode = False
activeslide.Shapes(3).Delete

''=======================================================================================================================
''EQUITY bubblechart MKt cap
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity - Capitalizzazione e Allocazione  "
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

bubble_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\bubblechart.emf"
Set bubble = activeslide.Shapes.AddPicture(bubble_path, msoFalse, msoTrue, 14.17, 54.432, 685.984, 447.02)

Application.CutCopyMode = False
activeslide.Shapes(2).Delete

''=======================================================================================================================
''EQUITY RISK ALLOCATION BAR
''=======================================================================================================================
'
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity - Capitalizzazione e Allocazione  "
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

bubble_path = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\Immagini update portafoglio\barchart_risk.emf"
Set bubble = activeslide.Shapes.AddPicture(bubble_path, msoFalse, msoTrue, 14.17, 54.432, 685.984, 447.02)

Application.CutCopyMode = False
activeslide.Shapes(2).Delete

''=======================================================================================================================
'''''ANALISI FI -SLIDE DIVISORIA
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Portafoglio Obbligazionario - Analisi"
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

With activeslide.Shapes(1)
    .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Left = 17
    .Top = 236
End With
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse
activeslide.Shapes(2).Delete


'=======================================================================================================================
'Rendimenti Fixed Income
'=======================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Rendimenti Fixed Income"
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

Set rfi = ws_tables.Range("DF3:DL43")
rfi.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 25.51
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 470.83
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 413.85

activeslide.Shapes(2).Delete
Application.CutCopyMode = False

''=======================================================================================================================
'''''BLOOMBERG-SLIDE DIVISORIA
''=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Portafoglio FERI PIR - Analisi Bloomberg"
    .Font.Size = 20
    .Font.Color = rgb(0, 0, 139)
    .Font.Name = "Georgia"
    .Font.Bold = True
End With

With activeslide.Shapes(1)
    .TextFrame.TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
    .Left = 17
    .Top = 236
End With
'Activeslide.Shapes(1).Range.Align msoAlignLefts, msoFalse
activeslide.Shapes(2).Delete

''=======================================================================================================================
'''''Pasting 10 bbl images
''=======================================================================================================================
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

AppActivate ("Microsoft PowerPoint")
Set activeslide = Nothing
Set newPPT = Nothing

ws_tables.Activate
Application.ScreenUpdating = True
MsgBox "FINITO!"

End Sub


