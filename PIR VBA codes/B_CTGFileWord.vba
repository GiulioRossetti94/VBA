'Sub CTG() ---> it produces the word file of the weekly CTG.
'
'Private Function getDatesAndMonth() ---> extract the day and month and year from Now() function in a format that is needed in the CTG word file
'
'Private Function checkFolderExistance() ---> checks if a folder exists, the input variable is the string address of the folder
'
'Private Function createFolder(ByVal strFolderPath As String) ---> INPUT= string of address. It checks whether a folder exists and if not, it creates each subfolder missing
'

Sub CTGMonday()
Call CTG
'incase the workbook is not closed, will set the timer for the next day
Application.OnTime TimeSerial(12, 0, 0), "CTGMonday"
End Sub

Sub CTG()
Dim wrdApp As Word.Application
Dim wrdDoc As Word.Document
Dim i As Integer
Dim ws As Excel.Worksheet

Set ws = Sheets("Tables")
nEQT = Sheets("Data").Range("E7") + 6
nETF = Sheets("Data").Range("F7") + 5
Set pETF = ws.Range(cells(3, 43), cells(nETF, 56))
Set pEqt = ws.Range(cells(3, 17), cells(nEQT, 32))
Set PortWeek = ws.Range(cells(3, 70), cells(38, 77))

Name_doc = "Verbale CTG FERI-PIR " & Format(Now(), "dd.mm.yyyy") & ".docx"
dt = Format(Now(), "mm.yy")
Y = Format(Now(), "yyyy")
nameSubFolder = Format(Now(), "yyyymmdd") & "_Asset allocation"
strFolder = "Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\" & Y & "\" & dt & "\" & nameSubFolder
createFolder (strFolder)


date_today = getDatesAndMonth()
day_today = date_today(0)
month_today = date_today(1)
year_today = date_today(2)


Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add(Template:="Y:\Mobiliare\08 Finint Economia Reale Italia\00_Documenti_Reportistica\03 Comitati\CTGTemplate.dotx", Newtemplate:=False, DocumentType:=0)

With wrdDoc
    .Content.InsertAfter "VERBALE DEL COMITATO TECNICO DI GESTIONE IN RELAZIONE " & _
                           " ALLA GESTIONE DEL FONDO " & Chr(34) & "FININT ECONOMIA REALE ITALIA " & Chr(34) & _
                           Chr(10) & day_today & " - " & day_today + 4 & " " & month_today & " " & year_today
    .Paragraphs(1).Range.Font.Size = 16
    .Paragraphs(1).Range.Font.Bold = True
    .Paragraphs(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Paragraphs(2).Range.Font.Size = 16
    .Paragraphs(2).Range.Font.Bold = True
    .Paragraphs(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter


    .Content.InsertAfter Chr(10) & Chr(10) & Chr(10)
    .Paragraphs(5).Range.Font.Size = 12
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Conegliano, " & day_today & "-" & day_today + 4 & " " & month_today & " " & year_today
    .Paragraphs(6).Range.Font.Size = 12
    .Paragraphs(6).Range.Font.Italic = True
    .Paragraphs(6).Range.Font.Bold = False
    .Paragraphs(6).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
    
    .Content.InsertAfter Chr(10) & Chr(10)
    .Content.InsertParagraphAfter
    
    .Content.InsertAfter "Punti all'ordine del giorno:" & Chr(10) & Chr(10)
    .Paragraphs(9).Range.Font.Italic = False
    
    inputDiscussion = InputBox("numero punti ordine del giorno", "Num", 1)

    For i = 1 To inputDiscussion
        If i = 1 Then
        pOrd = InputBox("Inserire punto" & i, , "Update di Portafoglio")
        ElseIf i = 2 Then
        pOrd = InputBox("Inserire punto" & i, , "Proposte di Asset Allocation")
        Else
        pOrd = InputBox("Inserire punto" & i)
        End If
        .Content.InsertAfter i & ")  " & pOrd
        .Content.InsertParagraphAfter
        .Paragraphs(10 + i).Range.Font.Italic = False
    Next
'    inputDiscussion = 2
    
'    For i = 1 To inputDiscussion
'        If i = 1 Then
'            pOrd = "Update di Portafoglio"
'        ElseIf i = 2 Then
'            pOrd = "Proposte di Asset Allocation"
'        End If
'        .Content.InsertAfter i & ")  " & pOrd
'        .Content.InsertParagraphAfter
'        .Paragraphs(10 + i).Range.Font.Italic = False
'    Next
   
   .Content.InsertAfter Chr(10) & Chr(10) & " Sono presenti i Signori: " & Chr(10) & Chr(10) & _
            "- Riccardo Igne" & Chr(9) & Chr(9) & "- Direttore Investimenti Obbligazionari" & Chr(10) & _
            "- Daniele Vadori" & Chr(9) & Chr(9) & "- Direttore Investimenti Azionari" & Chr(10) & _
            "- Filippo Napoletano" & Chr(9) & Chr(9) & "- Front Office " & Chr(10) & _
            "- Fausto Chino" & Chr(9) & Chr(9) & Chr(9) & "- Risk Manager" & Chr(10) & _
            "- Mattia Tormena" & Chr(9) & Chr(9) & "- Middle Office" & Chr(10) & _
            "- Thomas Beggio" & Chr(9) & Chr(9) & "- Risk Manager" & Chr(10) & Chr(10) & _
            "Sono assenti i Signori: " & Chr(10) & "---"
    l_par = .Paragraphs.Count + 1
    .Content.InsertParagraphAfter
    .Paragraphs(l_par).Range.InsertBreak Type:=wdPageBreak
    
    .Content.InsertParagraphAfter
    l_par = .Paragraphs.Count
    .Content.InsertAfter "Documentazione esaminata:"
    
    With .Paragraphs(l_par).Range.Font
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineSingle
    End With
    .Content.InsertParagraphAfter
    With .Paragraphs(l_par + 1).Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    .Content.InsertParagraphAfter
    l_par = .Paragraphs.Count
    .Content.InsertAfter "Trattazione:"
    
    With .Paragraphs(l_par).Range.Font
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineSingle
    End With
    .Content.InsertParagraphAfter
    With .Paragraphs(l_par + 1).Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    .Content.InsertAfter "Il gestore azionario Daniele Vadori riepiloga la performance " & _
                        "settimanale e la composizione del portafoglio azionario al " & Format(Now(), "dd/mm") & ":"
    .Content.InsertParagraphAfter
    l_par = .Paragraphs.Count
    
    
    PortWeek.Copy
    .Paragraphs(l_par).Range.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
                                            Placement:=wdInLine, DisplayAsIcon:=False
                                            
    .Content.InsertParagraphAfter
    pEqt.Copy
    .Paragraphs(l_par + 1).Range.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
                                            Placement:=wdInLine, DisplayAsIcon:=False
                                            
    Set Myshape = .InlineShapes(2).ConvertToShape
        With Myshape
            .WrapFormat.Type = wdWrapTopBottom
'            .WrapFormat.DistanceTop = 353.07
            .RelativeVerticalPosition = wdRelativeVerticalPositionPage
            .Top = 150
        End With
        
        
    .Content.InsertParagraphAfter
    pETF.Copy
    .Paragraphs(l_par + 2).Range.PasteSpecial link:=False, DataType:=wdPasteEnhancedMetafile, _
                                            Placement:=wdInLine, DisplayAsIcon:=False
    Excel.Application.CutCopyMode = False
    .Content.InsertParagraphAfter
    l_par = .Paragraphs.Count
    .Content.InsertAfter "Delibera"
    
    With .Paragraphs(l_par).Range.Font
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Non sono state eseguite operazioni" & Chr(10)
    With .Paragraphs(l_par + 1).Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Delibera"
    
    With .Paragraphs(l_par + 3).Range.Font
        .Bold = True
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Il presente CTG si riserva la facoltà di modificare il timing e/o le modalità degli acquisti dei titoli " & _
                            "qualora le condizioni di mercato e il contesto macroeconomico dovessero mutare." & vbCrLf & _
                            "Dopo approfondita discussione sul materiale presentato e sulle azioni proposte, il comitato delibera così come segue: " & vbCrLf

    With .Paragraphs(l_par + 4).Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    
    With .Paragraphs(l_par + 5).Range.Font
        .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With
    .Content.InsertParagraphAfter
'    .Paragraphs(l_par + 6).Range.ListFormat.ApplyListTemplate ListTemplate:= _
'            ListGalleries(wdBulletGallery).ListTemplates(1)
    .Content.InsertParagraphAfter
    .Paragraphs(l_par + 7).Range.InsertBreak Type:=wdPageBreak
            
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Firmato" & Chr(10) & Chr(10) & _
                    Chr(9) & "Riccardo Igne - Direttore degli Investimenti Obbligazionari" & Chr(10) & Chr(10) & _
                    Chr(9) & "Daniele Vadori - Direttore degli Investimenti Azionari" & Chr(10) & Chr(10) & _
                    Chr(9) & "Fausto Chino - Risk Manager" & Chr(10) & Chr(10)
   With .Range(.Paragraphs(l_par + 8).Range.Start, .Paragraphs(.Paragraphs.Count).Range.End).Font
         .Bold = False
        .Italic = False
        .Underline = wdUnderlineNone
    End With

    .SaveAs FileName:=strFolder & "\" & Name_doc
    
End With



End Sub

Private Function getDatesAndMonth() As Variant
Dim returnVal(3)
dt = Now()
dt_day = Format(dt, "dd")
dt_year = Format(dt, "yyyy")
mnth = Format(dt, "m")

If mnth = 1 Then
    dt_month = "Gennaio"
ElseIf mnth = 2 Then
    dt_month = "Febbraio"
ElseIf mnth = 3 Then
    dt_month = "Marzo"
ElseIf mnth = 4 Then
    dt_month = "Aprile"
ElseIf mnth = 5 Then
    dt_month = "Maggio"
ElseIf mnth = 6 Then
    dt_month = "Giugno"
ElseIf mnth = 7 Then
    dt_month = "Luglio"
ElseIf mnth = 8 Then
    dt_month = "Agosto"
ElseIf mnth = 9 Then
    dt_month = "Settembre"
ElseIf mnth = 10 Then
    dt_month = "Ottobre"
ElseIf mnth = 11 Then
    dt_month = "Novembre"
ElseIf mnth = 12 Then
    dt_month = "Dicembre"
Else: dt_month = "ERROR"
End If

returnVal(0) = dt_day
returnVal(1) = dt_month
returnVal(2) = dt_year
getDatesAndMonth = returnVal
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

