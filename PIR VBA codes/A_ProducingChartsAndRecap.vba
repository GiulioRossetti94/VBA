'SUBROUTINES
'insert_last_label_in_charts() ----> loop through all the charts in a specific sheet to show only the last label of a chart series
'
'doyouWannaFit() ----> loop through all the charts in a specific sheet and set the minimum y-axes value to either 0 or a value based on the minimum of the series
'
'pumpThatStats() ----> loop through all the charts in a specific sheet and set titles of charts. the strings for a title are in 2 different cells. the first cell include the part of the title with bigger font size and in bold, while the second one the string that goes on a separate line with a smaller font size
'
'WordWillTellUsApart() ----> subroutine not used in any command buttom. It is meant to create the "Daily market recap" file in word rather than in PPT
'
'DoYouWannaBeMyPPT() ----> it creates the PPT "Daily market recap" and saves it as a pdf to the Front office folder.
'
'WeDontNeedNoEduMail() ----> it calls the DoYouWannaBeMyPPT() sub, attaches the pdf and sends the email
'
'function RangetoHTML() ---> it prints an Excel range to the body of an outlook email


Sub insert_last_label_in_charts()
'guess it's self explanatory

Dim oChart As ChartObject
Dim mySeries As Series
Dim ws As Worksheet
Dim nseries As Integer
Dim npoints As Integer

'activate sheet where charts are
Set ws = Sheets("Tables")

npoints = Sheets("Data").Range("NG2").Value
'loop through charts
For Each oChart In ws.ChartObjects

    nseries = oChart.Chart.SeriesCollection.Count
    With oChart.Chart.SeriesCollection(1)
        .HasDataLabels = True
        .ApplyDataLabels (xlDataLabelsShowNone)
        .Points(npoints).ApplyDataLabels
        .DataLabels.Position = xlLabelPositionAbove
        .DataLabels.NumberFormat = "#,##0.00"
        .DataLabels.Font.Bold = True
        .DataLabels.Font.Size = 12
        .DataLabels.Font.Color = "blue"
    End With
    

    
Next

End Sub

Sub doyouWannaFit()

Dim oChart As ChartObject
Dim mySeries As Series
Dim mins(), x
Dim ws As Worksheet

Set ws = Worksheets("Tables")

If ThisWorkbook.Worksheets("Tables").CheckBox1.Value = True Then

    For Each oChart In ws.ChartObjects
        oChart.Chart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
        oChart.Chart.Axes(xlValue, xlPrimary).MinimumScaleIsAuto = True
    Next

Else
    For Each oChart In ws.ChartObjects
        min_val = Application.Min(oChart.Chart.SeriesCollection(1).Values)
        v_ax = min_val - 0.05 * Abs(min_val)
        
        oChart.Chart.Axes(xlValue, xlSecondary).MinimumScale = v_ax
        oChart.Chart.Axes(xlValue, xlPrimary).MinimumScale = 0
    Next
    
End If
End Sub

Sub pumpThatStats()

Dim jCol As Integer
Dim oChart As ChartObject
Dim ws1 As Worksheet
Dim ws2 As Worksheet

Set ws1 = Sheets("Data")
Set ws2 = Sheets("Tables")

jCol = 0

For Each oChart In ws2.ChartObjects

    str1 = ws1.cells(6, 374 + jCol)
    str2 = ws1.cells(4, 374 + jCol)
    
    len_1 = Len(str1)
    len_2 = Len(str2)
    
    With oChart.Chart
        .HasTitle = True
        .ChartTitle.Text = str1 & str2
        .ChartTitle.Characters(1, len_1).Font.Size = 18
        .ChartTitle.Characters(1 + len_1, len2).Font.Size = 12
        .ChartTitle.Characters(1 + len_1, len2).Font.Bold = False
    End With
    jCol = jCol + 2

Next
End Sub

Sub WordWillTellUsApart()
'create report in word

Dim wd As Object
Dim ObjDoc As Object
Dim FilePath As String
Dim FileName As String
Dim WordDoc As Word.Document
Dim ws As Worksheet
Dim nEQT As Integer

FilePath = "Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Daily Market Chart\Word\"

Application.ScreenUpdating = False

Set ws = Sheets("Tables")
On Error Resume Next
Set wd = GetObject(, "Word.Application")

If wd Is Nothing Then
    Set wd = CreateObject("Word.Application")
Else
    Set ObjDoc = wd.Documents(FileName)
    GoTo OpenAlready
End If
OpenAlready:
On Error GoTo 0

wd.Visible = True
wd.Documents.Add

'====================================================================================


Set Port1d = ws.Range(cells(3, 58), cells(53, 66))
Port1d.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Set PortWeek = ws.Range(cells(3, 70), cells(38, 77))
PortWeek.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

wd.Selection.InsertBreak Type:=wdSectionBreakNextPage
wd.Selection.PageSetup.Orientation = wdOrientLandscape

Set Summary = ws.Range("G4:N36")
Summary.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

nEQT = Sheets("Data").Range("E7") + 6
Set pEqt = ws.Range(cells(3, 18), cells(nEQT, 32))
pEqt.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

nETF = Sheets("Data").Range("F7") + 5
Set pETF = ws.Range(cells(3, 43), cells(nETF, 55))
pETF.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False
    
wd.Selection.InsertBreak Type:=wdSectionBreakNextPage
wd.Selection.PageSetup.Orientation = wdOrientPortrait

Set wei = ws.Range(cells(3, 36), cells(nEQT, 39))
wei.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False


Set ind = ws.Range(cells(3, 81), cells(59, 88))
ind.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Set top5 = ws.Range(cells(3, 93), cells(54, 107))
top5.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Set top5_2 = ws.Range(cells(54, 93), cells(85, 107))
top5_2.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Set bot5 = ws.Range(cells(87, 93), cells(137, 107))
bot5.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Set bot5_2 = ws.Range(cells(137, 93), cells(169, 107))
bot5_2.Copy
wd.Selection.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
    Placement:=wdInLine, DisplayAsIcon:=False

Application.ScreenUpdating = True
End Sub

Sub DoYouWannaBeMyPPT()

Dim newPPT As PowerPoint.Application
Dim Aslide As PowerPoint.Slide
Dim FileName As String
Dim ws_data As Worksheet
Dim ws_tables As Worksheet
Dim num_stocks, num_etf As Integer
Dim Mkt_sum, daily, weekly, industry, port_EQT, port_ETF, allocation As Range

FName = "\Daily Recap" & Format(Now, "dd.mm.yy") & ".pdf"
folderWMonthName = "Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Daily Market Chart\" & Format(Now(), "yyyy") & _
                    "\" & Format(Now(), "mm.yy")
createFolder (folderWMonthName)
FileName = folderWMonthName & FName


Set ws_data = Sheets("Data")
Set ws_tables = Sheets("Tables")

ws_tables.Activate
num_stocks = ws_data.Range("E7") + 6
num_etf = ws_data.Range("F7") + 5

Set Mkt_sum = ws_tables.Range("G4:N35")
Set daily = ws_tables.Range("BF4:BN53")
Set weekly = ws_tables.Range("BR4:BZ38")
Set industry = ws_tables.Range("CC4:CK41")
Set port_EQT = ws_tables.Range(cells(3, 17), cells(num_stocks, 32))
Set port_ETF = ws_tables.Range(cells(3, 43), cells(num_etf, 56))
Set allocation = ws_tables.Range(cells(3, 36), cells(num_stocks + num_etf, 41))

On Error Resume Next
    Set newPPT = GetObject(, "PowerPoint.Application")
On Error GoTo 0

If newPPT Is Nothing Then
    Set newPPT = New PowerPoint.Application
End If

If newPPT.Presentations.Count = 0 Then
    newPPT.Presentations.Add (msoCTrue)
End If

Set PPTReport = newPPT.Presentations.Open("C:\Users\bloomberg03\Desktop\Daily Market Chart\daily market_template.pptx")
Application.ScreenUpdating = False
newPPT.Visible = msoTrue


'======================================================================================================================================
'DAILY RECAP
'======================================================================================================================================
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

daily.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 211.46
newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.425
newPPT.ActiveWindow.Selection.ShapeRange.Width = 331.08
newPPT.ActiveWindow.Selection.ShapeRange.Height = 426.33
activeslide.Shapes(2).Delete

'======================================================================================================================================
'WEEKLY RECAP
'======================================================================================================================================

If ThisWorkbook.Worksheets("Tables").CheckBox2.Value = True Then
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
    
    activeslide.Shapes(1).Left = 17
    activeslide.Shapes(1).Top = 24
    With activeslide.Shapes(1).TextFrame.TextRange
        .Text = "Weekly Recap"
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
    
    weekly.CopyPicture
    activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 130.46
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.425
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 331.08
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 426.33
    activeslide.Shapes(2).Delete
End If

'======================================================================================================================================
'WEEKLY RECAP
'======================================================================================================================================

If ThisWorkbook.Worksheets("Tables").CheckBox3.Value = True Then
    newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
    newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
    Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
    
    activeslide.Shapes(1).Left = 17
    activeslide.Shapes(1).Top = 24
    With activeslide.Shapes(1).TextFrame.TextRange
        .Text = "Industry Recap"
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
    
    industry.CopyPicture
    activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 130.46
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.425
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 331.08
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 426.33
    activeslide.Shapes(2).Delete
End If

'======================================================================================================================================
'MARKET
'======================================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)
Set ind = ws_tables.Range("G4:N35")
ind.CopyPicture

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
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



activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 19
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 56.04
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 340
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 340
activeslide.Shapes(2).Delete


'======================================================================================================================================
'Equity Portfolio
'======================================================================================================================================

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

port_EQT.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 1.133
newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
newPPT.ActiveWindow.Selection.ShapeRange.Width = 552
newPPT.ActiveWindow.Selection.ShapeRange.Height = 397
activeslide.Shapes(2).Delete

'======================================================================================================================================
'ETF Portfolio
'======================================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Equity Portfolio - ETF"
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

port_ETF.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
newPPT.ActiveWindow.Selection.ShapeRange.Left = 1.133
newPPT.ActiveWindow.Selection.ShapeRange.Top = 72
newPPT.ActiveWindow.Selection.ShapeRange.Width = 700
newPPT.ActiveWindow.Selection.ShapeRange.Height = 87
activeslide.Shapes(2).Delete

'======================================================================================================================================
'Allocation
'======================================================================================================================================

newPPT.ActivePresentation.Slides.Add newPPT.ActivePresentation.Slides.Count + 1, ppLayoutText
newPPT.ActiveWindow.View.GotoSlide newPPT.ActivePresentation.Slides.Count
Set activeslide = newPPT.ActivePresentation.Slides(newPPT.ActivePresentation.Slides.Count)

activeslide.Shapes(1).Left = 17
activeslide.Shapes(1).Top = 24
With activeslide.Shapes(1).TextFrame.TextRange
    .Text = "Asset Allocation"
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

allocation.CopyPicture
activeslide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select

newPPT.ActiveWindow.Selection.ShapeRange.Left = 130.46
newPPT.ActiveWindow.Selection.ShapeRange.Top = 54.425
newPPT.ActiveWindow.Selection.ShapeRange.Width = 331.08
newPPT.ActiveWindow.Selection.ShapeRange.Height = 526.33
activeslide.Shapes(2).Delete

'======================================================================================================================================
'Charts best and worst performing
'======================================================================================================================================

Call doyouWannaFit
Call pumpThatStats
Call insert_last_label_in_charts

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
   ' ActiveSheet.Shapes(ActiveChart.Parent.Name).Line.Visible = msoFalse
    ActiveChart.ChartArea.Copy
    activeslide.Shapes.PasteSpecial(DataType:=5, link:=msoFalse).Select
    newPPT.ActiveWindow.Selection.ShapeRange.Left = 9.637795278
    newPPT.ActiveWindow.Selection.ShapeRange.Top = 105.44
    newPPT.ActiveWindow.Selection.ShapeRange.Width = 702.42
    newPPT.ActiveWindow.Selection.ShapeRange.Height = 226.77
    'ActiveSheet.Shapes(ActiveChart.Parent.Name).Line.Visible = msoTrue
    activeslide.Shapes(2).Delete
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01"))
    
Next i
Debug.Print FileName
If ThisWorkbook.Worksheets("Tables").CheckBox4.Value = True Then
    PPTReport.SaveAs FileName, ppSaveAsPDF
    PPTReport.Close
End If

End Sub

Sub WeDontNeedNoEduMail()
Application.ScreenUpdating = False

Dim rPIR As Range

'MAIL
Dim OutApp As Object
Dim OutMail As Object
Dim FileName As String

FName = "\Daily Recap" & Format(Now, "dd.mm.yy") & ".pdf"
folderWMonthName = "Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Daily Market Chart\" & Format(Now(), "yyyy") & _
                    "\" & Format(Now(), "mm.yy")

FileName = folderWMonthName & FName

Set rPIR = ThisWorkbook.Worksheets("Tables").Range("BF4:BN53")

ThisWorkbook.Worksheets("Tables").CheckBox4.Value = True
Call DoYouWannaBeMyPPT

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(olMailItem)

On Error Resume Next

With OutMail

    If ThisWorkbook.Worksheets("Tables").CheckBox6.Value = True Then
        emAdr = Sheets("Tables").Range("B27").Text
        .To = "giulio.rossetti@finint.com" & ";" & emAdr
    Else
        .To = "giulio.rossetti@finint.com"
    End If
    
    '.CC = "daniele.vadori@finint.com"
    .BCC = ""
    .Subject = "Daily Recap" & " " & Format(Now, "dd.mm.yyyy")
    .HTMLbody = RangetoHTML(rPIR)

    .Attachments.Add FileName
    .Display
    If ThisWorkbook.Worksheets("Tables").CheckBox5.Value = True Then .send
        
    

    '.Send > per l'invio automatico
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub



Function RangetoHTML(rng As Range)

    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .cells(1).PasteSpecial Paste:=8
        .cells(1).PasteSpecial xlPasteValues, , False, False
        .cells(1).PasteSpecial xlPasteFormats, , False, False
        .cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function



Private Function checkFolderExistance(ByVal strFolderPath As String) As Boolean

checkFolderExistance = Dir(strFolderPath, vbDirectory) <> vbNullString


End Function

Sub testingFunctions()

dt = Format(Now(), "mm.yy")
a = checkFolderExistance("Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\2018\" & dt)

MsgBox a
End Sub


Private Function createFolder(ByVal strFolderPath As String) As Boolean

Dim strCurrentFolder As String
Dim astrFolders() As String
Dim i As Integer

On Error GoTo ReturnFalse

If checkFolderExistance(strFolderPath) Then GoTo ReturnTrue
astrFolders = Split(strFolderPath, "\")

For i = LBound(astrFolders) To UBound(astrFolders)
  If astrFolders(i) <> vbNullString Then
      If strCurrentFolder <> vbNullString Then
        strCurrentFolder = strCurrentFolder & "\"
      End If
      strCurrentFolder = strCurrentFolder & astrFolders(i)
      If Not checkFolderExistance(strCurrentFolder) Then
        Call MkDir(strCurrentFolder)
      End If
    End If
  Next i
  
ReturnTrue:
  createFolder = True
  Exit Function
  
ReturnFalse:
  createFolder = False
    
End Function

