Sub DoFactsheet()
Dim wrdApp As Word.Application
Dim wrdDoc As Word.Document
Dim i As Integer
Dim ws As Excel.Worksheet
Dim objInLineShape As InlineShape
Dim objShape As Shape
Dim wdTable As Object
Dim oCell As Object

'Name_doc = "Verbale CTG Finint Dynamic - " & Format(Now(), "ddmmmmyy") & ".docx"

'day_td = CDbl(Format(Now(), "dd"))
'dt = Format(Now(), "mm.yyyy")
'y = Format(Now(), "yyyy")
'nWeek = day_td Mod 7

Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = True
Set wrdDoc = wrdApp.Documents.Add(Template:="Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\FStemplate1.dotx", Newtemplate:=False, DocumentType:=0)

With wrdDoc
    l_par = .Paragraphs.Count
    .Paragraphs(l_par).Range.InlineShapes.AddPicture ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\FS\perf_plt.jpg")

    For Each objInLineShape In .InlineShapes
        objInLineShape.ConvertToShape
    Next objInLineShape

    a = .Shapes.Count
    With .Shapes(a)
        .Top = -160
        .Left = 20
        .Height = 229.8898
        .Width = 315
          
    End With
    
 Set wdTable = .Tables(2)
 NR = wdTable.Rows.Count
 nc = wdTable.Columns.Count
  Debug.Print nc
 Set oCell = wdTable.cell(2, 1).Range
 oCell = "CANE"
 

End With

End Sub
