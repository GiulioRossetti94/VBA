'SCRIPT for "PORTFOLIO UPDATE" POWERPOINT PRESENTATION
'
'
'
'
'

Sub generateReport()

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

'=======================================================================================================================
''''SLIDE RENDIMENTI MERCATO'''''''''''''''''''''''''''''''''''''''''''''''''
'=======================================================================================================================
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
'=======================================================================================================================
'ANDAMENTO RACCOLTA
'=======================================================================================================================
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
    & vbCrLf & "Nel mese di Gennaio: " & Format(Range("R20").Text, "Currency") _
    & vbCrLf & "Nel mese di Febbraio: " & Format(Range("R21").Text, "Currency") _
    & vbCrLf & "Nel mese di Marzo: " & Format(Range("R22").Text, "Currency") _
    & vbCrLf & "Nel mese di Aprile: " & Format(Range("R23").Text, "Currency") _
    & vbCrLf & "Nel mese di Maggio: " & Format(Range("R24").Text, "Currency") _
    & vbCrLf & "Nel mese di Giugno: " & Format(Range("R25").Text, "Currency") _
    & vbCrLf & "Nel mese di Luglio: " & Format(Range("R26").Text, "Currency") _
    & vbCrLf & "Nel mese di Agosto: " & Format(Range("R27").Text, "Currency") _
    & vbCrLf & "Nel mese di Settembre: " & Format(Range("R28").Text, "Currency") _
    & vbCrLf & "Nel mese di Ottobre: " & Format(Range("R29").Text, "Currency") _
    & vbCrLf & "Nel mese di Novembre: " & Format(Range("R30").Text, "Currency") _
    & vbCrLf & "Nel mese di Dicembre: " & Format(Range("R31").Text, "Currency") _
               
    .Font.Size = 11
    .Font.Color = rgb(0, 0, 0)
    .Font.Name = "Georgia"
    .Font.Bold = False
End With
Application.CutCopyMode = False
'=======================================================================================================================
'Asset allocation
'=======================================================================================================================
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

Set Nav = Sheets("Asset Allocation").Range("A3:C11")
Nav.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 31.18
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 85.03
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 198
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 100
    
Set esp = Sheets("Asset Allocation").Range("A11:E30")
esp.CopyPicture

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 31.18
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 250
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 381.82
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 216.28
    
Set tas = Sheets("Performance").Range("AG1:AL15")
tas.CopyPicture

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 500
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 250
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 180
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 154.5
Application.CutCopyMode = False
'=======================================================================================================================
'Asset allocation 2/2
'=======================================================================================================================

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

Set pie = Sheets("Asset Allocation").Range("E1:L14")
pie.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = -65.2
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
'=======================================================================================================================
'Portfolio 1
'=======================================================================================================================

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

'Sheets("Asset Allocation").Activate
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
Set Equity = Sheets("Asset Allocation").Range("A37:J72")
Equity.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
'    newPPT.ActiveWindow.Selection.ShapeRange.LockAspectRatio = msoFalse
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 331.37

'lRow2 = Cells(118, 1).End(xlUp).Offset(1, 0).Row
'Set etf = Sheets("Asset Allocation").Range(Cells(118, "A"), Cells(lRow2, "I"))
Set etf = Sheets("Asset Allocation").Range("A119:J125")
etf.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 400
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 61

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'Portfolio 2
'=======================================================================================================================

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

Set fi_ = Sheets("Asset Allocation").Range("A79:J113")

fi_.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 331.37

Set gvt = Sheets("Asset Allocation").Range("A130:J132")
gvt.CopyPicture
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))

activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 53
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 400
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 650.56
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 32

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'Rendimenti Equity
'=======================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Rendimenti Equity"
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


'Sheets("Rnd Daily").Activate

Set re = ws_tables.Range("Q3:AG36")

re.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 25.51
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 676.06
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 374.455

activeslide.Shapes(2).Delete

Call generateReportPythonTest

Application.CutCopyMode = False
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

'Sheets("Rnd Daily").Activate
Set rfi = ws_tables.Range("DF3:DL38")
rfi.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 25.51
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 470.83
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 413.85

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'Rendimenti ETF
'=======================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Rendimenti ETF"
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


Set rETF = ws_tables.Range("AQ3:BD9")
rETF.CopyPicture
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 11.622
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.42
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 708
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 85

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'Rendimenti Indici
'=======================================================================================================================

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

'Sheets("Indici Italia").Activate
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
'=======================================================================================================================
'Add sheets to paste bloomberg pics
'=======================================================================================================================

Name_to_check = "REPORT_creazione"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True
        
    End If
Next Sheet

If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "REPORT_creazione"
End If
    
'=======================================================================================================================
'IMMAGINI BLOOMBERG
'=======================================================================================================================
DoEvents
    Application.Wait (Now + TimeValue("0:00:001"))
Sheets("REPORT_creazione").Activate
Dim strFolder As String
Dim strFileName As String
Dim objPic As Picture
Dim rngCell As Range

strFolder = "C:\Users\bloomberg03\Desktop\BBL_pic"
If Right(strFolder, 1) <> "\" Then
    strFolder = strFolder & "\"
End If

With Worksheets("REPORT_creazione")
    .Columns(1).ColumnWidth = 106
    .Rows.RowHeight = 400
End With

Set rngCell = Range("A1")

strFileName = Dir(strFolder & "*.jpg", vbNormal)

'Do While Len(strFileName) > 0
    For i = 1 To 10
    strFileName = Dir(strFolder & i & "*.jpg", vbNormal)
    Set objPic = ActiveSheet.Pictures.Insert(strFolder & strFileName)
        With objPic
            .Left = rngCell.Left
            .Top = rngCell.Top
            .Height = rngCell.RowHeight
            .Placement = xlMoveAndSize
        End With
    Debug.Print strFileName
    Set rngCell = rngCell.Offset(1, 0)
    strFileName = Dir(strFolder & strFileName)
    Next
'Loop
'=======================================================================================================================
'IMMAGINI BLOOMBERG 1
'=======================================================================================================================
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


Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A1")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False

'=======================================================================================================================
'IMMAGINI BLOOMBERG 1
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Caratteristiche portafoglio"
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


Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A2")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 2
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Flussi di cassa"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A3")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 3
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Tassi chiave"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A4")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 4
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Volatilità"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A5")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
   newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 5
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Comparazione VaR (P&L)"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A6")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 6
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Comparazione VaR (rend%)"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A7")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 7
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Distribuzione VaR"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A8")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 8
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Peggiori scenari Fixed Income"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A9")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete
Application.CutCopyMode = False
'=======================================================================================================================
'IMMAGINI BLOOMBERG 9
'=======================================================================================================================
newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

 activeslide.Shapes(1).Left = 17
 activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Peggiori scenari Equity"
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

Sheets("REPORT_creazione").Activate
Set cp = Sheets("REPORT_creazione").Range("A10")
cp.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 36.85
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 397

activeslide.Shapes(2).Delete

Sheets("REPORT_creazione").Activate
Application.DisplayAlerts = False
Worksheets("REPORT_creazione").Delete



AppActivate ("Microsoft PowerPoint")
Set activeslide = Nothing
Set newPPT = Nothing



ws_tables.Activate
Application.ScreenUpdating = True




End Sub