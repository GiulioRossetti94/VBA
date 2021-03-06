Sub Proposta()

Application.ScreenUpdating = False

'CREA COPIA DEI FOGLI DA INSERIRE NELLA PROPOSTA A.A.
Dim myFogli As Variant
myFogli = Array("Asset Allocation", "Simulazione A.A.", "PTF ETF", "PTF EQUITY", "PTF BOND")
Sheets(myFogli).Copy

'COPIA E INCOLLA VALORI
Sheets("Asset Allocation").Select
Range("A1", "Z150").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

Sheets("Simulazione A.A.").Select
Range("A1", "Z150").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

Sheets("PTF ETF").Select
Range("A1", "BJ50").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

Sheets("PTF EQUITY").Select
Range("A1", "BJ50").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

Sheets("PTF BOND").Select
Range("A1", "BJ50").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues
ActiveSheet.Paste
Application.CutCopyMode = False
Range("A1").Select

Sheets("Simulazione A.A.").Select
Range("A1").Select


'SALVATAGGIO FILE
nomefile = "Simulazione_" & Format(Now, "ddmmyy") & ".xlsx"
ActiveSheet.SaveAs FileName:=("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Simulazioni A.A\" & nomefile)


'MAIL
Dim OutApp As Object
Dim OutMail As Object

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

On Error Resume Next
   
With OutMail
    .To = "daniele.vadori@finint.com"
    .CC = ""
    .BCC = ""
    .Subject = "Proposta FERI-PIR" & " " & Format(Now, "dd.mm.yyyy")
    .Body = "Ciao," & vbCrLf & "in allegato il file Excel con le proposte d'investimento per il fondo Finint Economia Reale Italia." & vbCrLf & " " & vbCrLf & ""
    '.Attachments.Add ActiveWorkbook.FullName
    .Attachments.Add ("Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Simulazioni A.A\" & nomefile)
    .Display
    '.Send > per l'invio automatico
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
    

'CHIUDI
MsgBox ("File salvato")
ActiveWorkbook.Close (True)



End Sub