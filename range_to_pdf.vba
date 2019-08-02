Sub to_pdf()
Application.ScreenUpdating = False
Dim wd As Object
Dim ObjDoc As Object
Dim FilePath As String
Dim FileName As String
Dim rng As Range
FilePath = "Y:\Temp\Stefano\file\"

Dim WordDoc As Word.Document

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

Name_file = Range("A1")
range_cell = Range("A2")
Set rng = Range(range_cell)
rng.Copy
wd.Selection.PasteSpecial Link:=False, DataType:=wdPasteEnhancedMetafile, _
Placement:=wdInLine, DisplayAsIcon:=False


Application.CutCopyMode = False
wd.ActiveDocument.ExportAsFixedFormat FilePath & Name_file & "_" & Format(Now, "dd.mm.yy") & ".pdf", 17, OpenAfterExport:=True
wd.ActiveDocument.SaveAs FileName:=FilePath & Name_file & "_" & Format(Now, "dd.mm.yy") & ".docx"
wd.Quit
Application.ScreenUpdating = True
Call email_pdf
End Sub

'=========================================================================================================

Sub email_pdf()

Dim OutApp As Object
Dim OutMail As Object
Dim FileName As String

FilePath = "Y:\Temp\Stefano\file\"   'output path
Name_file = Range("A1")
FileName = FilePath & Name_file & "_" & Format(Now, "dd.mm.yy") & ".pdf"

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(olMailItem)



On Error Resume Next
str1 = "PDF,"
With OutMail

    .To = "giulio.rossetti@.com"

    .BCC = ""
    .body = str1
    .Subject = Name_file & " " & Format(Now, "dd.mm.yyyy")
        

    .Attachments.Add FileName
    .Display
'    .Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing





End Sub
