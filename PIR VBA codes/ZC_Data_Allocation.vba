Sub PreparingDataAllocation()
Dim ws_d As Worksheet

Dim rng_eqt As Range
Dim rng_etf As Range
Dim rng_fi As Range
Dim rng_gov As Range
Dim rng_port As Range

Dim eqt() As Variant
Dim etf() As Variant
Dim fi() As Variant
Dim gov() As Variant
Dim listed() As Variant

Dim unlisted As New Collection

Dim port() As Variant
Dim charac() As Variant

Dim price() As Variant
Dim industry() As Variant
Dim indRng As Range
Dim retRng As Range
Dim namesEQT() As Variant
Dim nmRng As Range

'application.Calculation = xlManual

Set ws_d = Sheets("Asset Allocation")

n_stock = ws_d.cells(36, 1)
n_etf = ws_d.cells(36, 12)
n_fi = ws_d.cells(36, 23)
n_gov = ws_d.cells(36, 34)

ReDim eqt(1 To n_stock, 1 To 10)
ReDim etf(1 To n_etf, 1 To 10)
ReDim fi(1 To n_fi, 1 To 10)
ReDim gov(1 To n_gov, 1 To 10)
ReDim listed(1 To n_fi, 1 To 1)

Set rng_eqt = ws_d.Range(ws_d.cells(39, 1), ws_d.cells(39 + n_stock - 1, 10))
Set rng_etf = ws_d.Range(ws_d.cells(39, 12), ws_d.cells(39 + n_etf - 1, 21))
Set rng_fi = ws_d.Range(ws_d.cells(39, 23), ws_d.cells(39 + n_fi - 1, 32))
Set rng_gov = ws_d.Range(ws_d.cells(39, 34), ws_d.cells(39 + n_gov - 1, 43))

eqt = rng_eqt.Value
etf = rng_etf.Value
fi = rng_fi.Value
gov = rng_gov.Value
'Debug.Print n_stock; n_etf; n_fi; n_gov

'check if sheet(allocation) exists, if not, addd it
Name_to_check = "allocation"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True
    End If
Next Sheet
If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count - 1))
    ws.Name = "allocation"
End If

Set ws_f = Sheets("allocation")

Application.ScreenUpdating = False
Sheets("allocation").UsedRange.Delete
Application.ScreenUpdating = True

For k = 1 To n_fi
    If GetFontColor(ws_d.cells(38 + k, 24)) <> 1 Then unlisted.Add ws_d.cells(38 + k, 24)
Next k

For i = 1 To unlisted.Count
    Debug.Print unlisted(i)
Next i
'paste date in sheet

ws_f.Range(ws_f.cells(1, 1), ws_f.cells(1, UBound(eqt, 2))) = ws_d.Range(ws_d.cells(38, 23), ws_d.cells(38, 32)).Value
ws_f.Range(ws_f.cells(2, 1), ws_f.cells(n_stock + 1, UBound(eqt, 2))) = eqt
ws_f.Range(ws_f.cells(2 + n_stock, 1), ws_f.cells(1 + n_stock + n_etf, UBound(etf, 2))) = etf
ws_f.Range(ws_f.cells(2 + n_stock + n_etf, 1), ws_f.cells(1 + n_stock + n_etf + n_fi, UBound(fi, 2))) = fi
ws_f.Range(ws_f.cells(2 + n_stock + n_etf + n_fi, 1), ws_f.cells(1 + n_stock + n_etf + n_fi + n_gov, UBound(gov, 2))) = gov
ws_f.Range(ws_f.cells(2 + n_stock + n_etf + n_fi + n_gov, 1), ws_f.cells(1 + n_stock + n_etf + n_fi + n_gov + 1, UBound(gov, 2) - 1)) = "CASH"
ws_f.cells(2 + n_stock + n_etf + n_fi + n_gov, UBound(gov, 2) - 1) = ws_d.cells(9, 2).Value
ws_f.cells(2 + n_stock + n_etf + n_fi + n_gov, UBound(gov, 2)) = ws_d.cells(9, 2).Value / ws_d.cells(3, 2).Value

n_row = 1 + n_stock + n_etf + n_fi + n_gov
date_YTD = DateValue("12/28/2018")
date_MTD = Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(Now(), -1) + 1, -1)
date_W = Application.WorksheetFunction.WorkDay(Now(), -5)
date_D = Application.WorksheetFunction.WorkDay(Now(), -1)

ReDim port(1 To n_row, 1 To 10)

Set rng_port = ws_f.Range(ws_f.cells(2, 1), ws_f.cells(1 + n_stock + n_etf + n_fi + n_gov + 1, UBound(gov, 2)))
port = rng_port.Value

ReDim charac(1 To n_row, 1 To 18)

'~~~~~~~SET HEADERS~~~~~~~
ws_f.cells(1, 11).Value = "LISTED"
ws_f.cells(1, 12).Value = "CPN_TYP"
ws_f.cells(1, 13).Value = "CPN"
ws_f.cells(1, 14).Value = "MTY_YEARS_TDY"
ws_f.cells(1, 15).Value = "DUR_ADJ_MTY_BID"
ws_f.cells(1, 16).Value = Format(date_YTD, "yyyymmdd")
ws_f.cells(1, 17).Value = Format(date_MTD, "yyyymmdd")
ws_f.cells(1, 18).Value = Format(date_W, "yyyymmdd")
ws_f.cells(1, 19).Value = Format(date_D, "yyyymmdd")
ws_f.cells(1, 20).Value = "Equity Market"
ws_f.cells(1, 21).Value = "Capitalization (Mln)"
ws_f.cells(1, 22).Value = "Div yield"
ws_f.cells(1, 23).Value = "Target Price"
ws_f.cells(1, 24).Value = "Country"
ws_f.cells(1, 25).Value = "P/B"
ws_f.cells(1, 26).Value = "D/E"
ws_f.cells(1, 27).Value = "Prev CPN dt"
ws_f.cells(1, 28).Value = "Next CPN dt"
'ws_f.cells(1, 29).Value = "Dvd_ex_dt"
'~~~~~~~SET HEADERS~~~~~~~

For i = 1 To n_row

    For q = 1 To unlisted.Count
        If port(i, 2) = unlisted(q) Then
            charac(i, 1) = 0
            GoTo nit
'            Debug.Print port(i, 2)
        Else
            charac(i, 1) = 1
'            Debug.Print port(i, 2)
        End If
    Next q
nit:
    If port(i, 1) = "EQUITY" Or port(i, 1) = "CASH" Then
        charac(i, 2) = "nan"
        charac(i, 3) = "nan"
        charac(i, 4) = "nan"
        charac(i, 5) = "nan"
        charac(i, 17) = "nan"
        charac(i, 18) = "nan"
        
        price_YTD = Application.Index(Sheets("Data").Range("JF8:NA1000"), Application.Match(CDbl(date_YTD), Sheets("Data").Range("JE8:JE1000"), 0), Application.Match(ws_f.cells(1 + i, 2), Sheets("Data").Range("JF7:NA7"), 0))
        price_MTD = Application.Index(Sheets("Data").Range("JF8:NA1000"), Application.Match(CDbl(date_MTD), Sheets("Data").Range("JE8:JE1000"), 0), Application.Match(ws_f.cells(1 + i, 2), Sheets("Data").Range("JF7:NA7"), 0))
        price_W = Application.Index(Sheets("Data").Range("JF8:NA1000"), Application.Match(CDbl(date_W), Sheets("Data").Range("JE8:JE1000"), 0), Application.Match(ws_f.cells(1 + i, 2), Sheets("Data").Range("JF7:NA7"), 0))
        price_D = Application.Index(Sheets("Data").Range("JF8:NA1000"), Application.Match(CDbl(date_D), Sheets("Data").Range("JE8:JE1000"), 0), Application.Match(ws_f.cells(1 + i, 2), Sheets("Data").Range("JF7:NA7"), 0))
        MKT = Application.Index(Sheets("Monitor Azioni").Range("I:I"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        cap = Application.Index(Sheets("Monitor Azioni").Range("J:J"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        Div_yield = Application.Index(Sheets("Monitor Azioni").Range("AT:AT"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        tg_price = Application.Index(Sheets("Monitor Azioni").Range("X:X"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        PB = Application.Index(Sheets("Monitor Azioni").Range("BD:BD"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        DE = Application.Index(Sheets("Monitor Azioni").Range("BQ:BQ"), Application.Match(port(i, 2), Sheets("Monitor Azioni").Range("C:C"), 0))
        'ex_dt = "=IF(OR(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 29) & Chr(34) & ")=" & Chr(34) & "#N/A Field Not Applicable" & Chr(34) & ",BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 29) & Chr(34) & ")=" & Chr(34) & "#N/A Invalid Security" & Chr(34) & ",BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 29) & Chr(34) & ")=" & Chr(34) & "#N/A N/A" & Chr(34) & "),""nan"",BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 29) & Chr(34) & "))"
            

            
        If IsError(price_YTD) Or VarType(price_YTD) = vbString Then price_YTD = "nan"
        If IsError(price_MTD) Or VarType(price_MTD) = vbString Then price_MTD = "nan"
        If IsError(price_W) Or VarType(price_W) = vbString Then price_W = "nan"
        If IsError(price_D) Or VarType(price_D) = vbString Then price_D = "nan"
        If IsError(MKT) Then MKT = "nan"
        If IsError(cap) Then cap = "nan"
        If IsError(Div_yield) Or VarType(Div_yield) = vbString Then Div_yield = "nan"
        If IsError(tg_price) Or VarType(tg_price) = vbString Then tg_price = "nan"
        If IsError(PB) Or VarType(PB) = vbString Then PB = "nan"
        If IsError(DE) Or VarType(DE) = vbString Then DE = "nan"
        
    ElseIf InStr(port(i, 1), "FIXED INCOME") > 0 Then
        price_YTD = Application.Index(Sheets("STORICO PREZZI_FI").Range("F:F"), Application.Match(Format(date_YTD, "yyyymmdd") & port(i, 3), Sheets("STORICO PREZZI_FI").Range("E:E"), 0))
        price_MTD = Application.Index(Sheets("STORICO PREZZI_FI").Range("F:F"), Application.Match(Format(date_MTD, "yyyymmdd") & port(i, 3), Sheets("STORICO PREZZI_FI").Range("E:E"), 0))
        price_W = Application.Index(Sheets("STORICO PREZZI_FI").Range("F:F"), Application.Match(Format(date_W, "yyyymmdd") & port(i, 3), Sheets("STORICO PREZZI_FI").Range("E:E"), 0))
        price_D = Application.Index(Sheets("STORICO PREZZI_FI").Range("F:F"), Application.Match(Format(date_D, "yyyymmdd") & port(i, 3), Sheets("STORICO PREZZI_FI").Range("E:E"), 0))
        MKT = "nan"
        cap = "nan"
        Div_yield = "nan"
        tg_price = "nan"
'        PB = "=BDP(" & Chr(34) & port(i, 2) & Chr(34) & ", " & Chr(34) & "PREV_CPN_DT" & Chr(34) & ")"
'        DE = "=BDP(" & Chr(34) & port(i, 2) & Chr(34) & ", " & Chr(34) & "NXT_CPN_DT" & Chr(34) & ")"
        PB = "nan"
        DE = "nan"
        'ex_dt = "nan"
        If IsError(price_YTD) Or VarType(price_YTD) = vbString Then price_YTD = "nan"
        If IsError(price_MTD) Or VarType(price_MTD) = vbString Then price_MTD = "nan"
        If IsError(price_W) Or VarType(price_W) = vbString Then price_W = "nan"
        If IsError(price_D) Or VarType(price_D) = vbString Then price_D = "nan"
        
        
        charac(i, 2) = "=BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 12) & Chr(34) & ")"
        charac(i, 3) = "=IF(ISNUMBER(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 13) & Chr(34) & ")),BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 13) & Chr(34) & "),""nan"")"
        charac(i, 4) = "=IF(ISNUMBER(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 14) & Chr(34) & ")),BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 14) & Chr(34) & "),""nan"")"
        charac(i, 5) = "=IF(ISNUMBER(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 15) & Chr(34) & ")),BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 15) & Chr(34) & "),""nan"")"
        charac(i, 17) = "=BDP(" & Chr(34) & port(i, 2) & Chr(34) & ", " & Chr(34) & "PREV_CPN_DT" & Chr(34) & ")"
        charac(i, 18) = "=BDP(" & Chr(34) & port(i, 2) & Chr(34) & ", " & Chr(34) & "NXT_CPN_DT" & Chr(34) & ")"
        
    End If

        charac(i, 6) = price_YTD
        charac(i, 7) = price_MTD
        charac(i, 8) = price_W
        charac(i, 9) = price_D
        charac(i, 10) = MKT
        charac(i, 11) = cap
        charac(i, 12) = Div_yield
        charac(i, 13) = tg_price
        charac(i, 14) = "=IF(OR(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 24) & Chr(34) & ")=" & Chr(34) & "#N/A Field Not Applicable" & Chr(34) & ",BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 24) & Chr(34) & ")=" & Chr(34) & "#N/A Invalid Security" & Chr(34) & "),""nan"",BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.cells(1, 24) & Chr(34) & "))"
        charac(i, 15) = PB
        charac(i, 16) = DE
        'charac(i, 19) = ex_dt
'    Debug.Print charac(i, 5); charac(i, 6); charac(i, 7)
Next i



ws_f.Range(ws_f.cells(2, UBound(eqt, 2) + 1), ws_f.cells(2 + n_stock + n_etf + n_fi + n_gov, UBound(eqt, 2) + UBound(charac, 2))) = charac
   
 
'For i = (2 + n_stock + n_etf) To n_stock + n_etf + n_fi + n_gov + 1
'
'    ws_f.Cells(i, 12).Value = "=IF(ISNUMBER(BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.Cells(1, 12) & Chr(34) & ")),BDP(" & Chr(34) & port(i, 2) & Chr(34) & "," & Chr(34) & ws_f.Cells(1, 12) & Chr(34) & "),""nan"")"
'    ws_f.Cells(i, 12).Value = "=BDP(" & Chr(34) & port(i - 1, 2) & Chr(34) & "," & Chr(34) & ws_f.Cells(1, 13) & Chr(34) & ")"
'    ws_f.Cells(i, 12).Value = "=BDP(" & Chr(34) & port(i - 1, 2) & Chr(34) & "," & Chr(34) & ws_f.Cells(1, 14) & Chr(34) & ")"
'    ws_f.Cells(i, 12).Value = "=BDP(" & Chr(34) & port(i - 1, 2) & Chr(34) & "," & Chr(34) & ws_f.Cells(1, 15) & Chr(34) & ")"
'
'Next i
'
            
'    MKT(i, 1) = Application.Index(ws_m.Range(ws_m.Cells(7, 10), ws_m.Cells(7, 10).End(xlDown)), Application.Match(data(i + 1, 1), _
'                ws_m.Range(ws_m.Cells(7, 3), ws_m.Cells(7, 3).End(xlDown)), 0))
        


'ws_f.Range(ws_f.Cells(2, UBound(data, 2) + 1), ws_f.Cells(nEQT + 1, UBound(data, 2) + 1)) = MKT
'ws_f.Range(ws_f.Cells(1, UBound(data, 2) + 2), ws_f.Cells(nEQT + 1, UBound(data, 2) + 2)) = industry
'ws_f.Cells(1, UBound(data, 2) + 1) = "MKT_STOCKS"
''ws_f.Range(ws_f.Cells(2, UBound(data, 2) + 4), ws_f.Cells(goDown + 2, UBound(data, 2) + 4 + nEQT + 2)) = ret
''ws_f.Range(ws_f.Cells(1, UBound(data, 2) + 5), ws_f.Cells(1, UBound(data, 2) + 4 + nEQT + 2)) = namesEQT
''ws_f.Cells(1, UBound(data, 2) + 4) = "Date"
'
'ws_f.Range(ws_f.Cells(1, UBound(data, 2) + 4), ws_f.Cells(goDown + 1, UBound(data, 2) + 4 + nEQT + 20 + 2)) = price
'
'ws_f.Cells.NumberFormat = "General"
'ws_f.Range("AI:AI").NumberFormat = "yyyymmdd"
'Application.Calculation = xlCalculationAutomatic
'
'Application.DisplayAlerts = False
'strFullname = "C:\Users\bloomberg03\Desktop\PythonScript\PTF PIR.csv"
'ThisWorkbook.Sheets("JL Data").Copy
'ActiveWorkbook.SaveAs FileName:=strFullname, FileFormat:=xlCSV, CreateBackup:=False
'ActiveWorkbook.Close
'
'Application.DisplayAlerts = True


End Sub

Function GetFontColor(ByVal Target As Range) As Integer
    GetFontColor = Target.Font.ColorIndex
End Function