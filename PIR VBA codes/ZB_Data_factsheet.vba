Sub table_perf()
Dim ws As Worksheet
Dim d() As Variant

Application.Calculation = xlManual

Set ws = Sheets("Performance")
Set N = ws.cells(1, 8)
Set data = ws.Range(ws.cells(1, 9), ws.cells(N, 16))
ReDim d(1 To N, 1 To 8)

d = data.Value


Name_to_check = "data_perf"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True
    End If
Next Sheet
If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count - 1))
    ws.Name = Name_to_check
End If

Set ws_fact = Sheets(Name_to_check)

ws_fact.Range(ws_fact.cells(1, 1), ws_fact.cells(N, 8)) = d
ws_fact.Range(ws_fact.cells(1, 1), ws_fact.cells(N, 8)).NumberFormat = "General"

ws_fact.Range("A:A").NumberFormat = "yyyymmdd"

Application.DisplayAlerts = False
strFullname = "C:\Users\bloomberg03\Desktop\PythonScript\Factsheet\performance_pir_data"
ThisWorkbook.Sheets(Name_to_check).Copy
ActiveWorkbook.SaveAs FileName:=strFullname, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.Close

Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
End Sub