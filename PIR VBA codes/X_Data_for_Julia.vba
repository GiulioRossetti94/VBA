Sub PreparingData()
Dim ws_d As Worksheet
Dim ws_m As Worksheet
Dim rng As Range
Dim data() As Variant
Dim MKT() As Variant
Dim Ret() As Variant
Dim price() As Variant
Dim industry() As Variant
Dim indRng As Range
Dim retRng As Range
Dim namesEQT() As Variant
Dim nmRng As Range

Application.Calculation = xlManual


Set ws_d = Sheets("Data")
Set ws_m = Sheets("Monitor Azioni")

nEQT = ws_d.cells(7, 5)
nETF = ws_d.cells(7, 6)

goDown = ws_d.cells(7, 59)

Set rng = ws_d.Range(ws_d.cells(8, 3), ws_d.cells(8 + nEQT + nETF, 3).End(xlToRight))
Set retRng = ws_d.Range(ws_d.cells(9, 60), ws_d.cells(9 + goDown, 60 + nEQT + 2))
Set nmRng = ws_d.Range(ws_d.cells(7, 61), ws_d.cells(7, 61 + nEQT - 1 + 2))
Set indRng = ws_d.Range(ws_d.cells(8, 39), ws_d.cells(8 + nEQT, 39))
Set priceRng = ws_d.Range(ws_d.cells(7, 265), ws_d.cells(7 + goDown, 265 + nEQT + 20))


industry = indRng.Value
Ret = retRng.Value
price = priceRng.Value
data = rng.Value
namesEQT = nmRng.Value

ReDim MKT(1 To nEQT, 1 To 1)

For i = 1 To nEQT
    Err.Clear
    MKT(i, 1) = Application.Index(ws_m.Range(ws_m.cells(7, 10), ws_m.cells(7, 10).End(xlDown)), Application.Match(data(i + 1, 1), _
                ws_m.Range(ws_m.cells(7, 3), ws_m.cells(7, 3).End(xlDown)), 0))
Next i

Name_to_check = "JL Data"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True
    End If
Next Sheet
If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count - 1))
    ws.Name = "JL Data"
End If

Set ws_f = Sheets("JL Data")

Application.ScreenUpdating = False
Sheets("JL Data").UsedRange.Delete
Application.ScreenUpdating = True

ws_f.Range(ws_f.cells(1, 1), ws_f.cells(nEQT + 1 + nETF, UBound(data, 2))) = data
ws_f.Range(ws_f.cells(2, UBound(data, 2) + 1), ws_f.cells(nEQT + 1, UBound(data, 2) + 1)) = MKT
ws_f.Range(ws_f.cells(1, UBound(data, 2) + 2), ws_f.cells(nEQT + 1, UBound(data, 2) + 2)) = industry
ws_f.cells(1, UBound(data, 2) + 1) = "MKT_STOCKS"
'ws_f.Range(ws_f.Cells(2, UBound(data, 2) + 4), ws_f.Cells(goDown + 2, UBound(data, 2) + 4 + nEQT + 2)) = ret
'ws_f.Range(ws_f.Cells(1, UBound(data, 2) + 5), ws_f.Cells(1, UBound(data, 2) + 4 + nEQT + 2)) = namesEQT
'ws_f.Cells(1, UBound(data, 2) + 4) = "Date"

ws_f.Range(ws_f.cells(1, UBound(data, 2) + 4), ws_f.cells(goDown + 1, UBound(data, 2) + 4 + nEQT + 20 + 2)) = price

ws_f.cells.NumberFormat = "General"
ws_f.Range("AI:AI").NumberFormat = "yyyymmdd"
Application.Calculation = xlCalculationAutomatic

Application.DisplayAlerts = False
strFullname = "C:\Users\bloomberg03\Desktop\PythonScript\PTF PIR.csv"
ThisWorkbook.Sheets("JL Data").Copy
ActiveWorkbook.SaveAs FileName:=strFullname, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.Close

Application.DisplayAlerts = True


End Sub
