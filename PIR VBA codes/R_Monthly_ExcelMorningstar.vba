Sub MorningstarExcel()
Dim fbank As Worksheet
Dim data() As Variant

date_report = Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(Now(), -1) + 1, -1)

DateToGO = InputBox("inserisci data", , Format(date_report, "dd/mm/yyyy"))

nameFolder = Format(DateToGO, "yyyymmdd")
dateNameFile = Format(DateToGO, "dd.mm.yyyy")

year_file = Format(DateToGO, "yyyy")
month_file = Format(DateToGO, "mm yyyy")
month_filefolder = Format(DateToGO, "mm.yy")
strFolder = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Banca Finint\Dati portafoglio\" & year_file & "\" & month_filefolder
nameFiletoOpenBanca = "Fondo FERI - PIR " & Format(DateToGO, "mm.yy") & " VBA Formule.xlsx"

fileExcelBanca = strFolder & "\" & nameFiletoOpenBanca

'working WITH banca finint excel

Close #1
    Workbooks.Open fileExcelBanca
Set fbank = Workbooks(nameFiletoOpenBanca).Sheets("Composizione PTF Fondo")

N = fbank.Range(fbank.cells(3, 2), fbank.cells(3, 2).End(xlDown)).Count - 2

DQ = Chr(34)
ReDim data(1 To N, 1 To 11)

Neq = 0
For i = 1 To N
    ticker = Replace(Replace(fbank.cells(2 + i, 2).Formula, "=BDP(" & DQ, ""), DQ & "," & DQ & "NAME" & DQ & ")", "")
    ticker = Replace(ticker, DQ & "," & DQ & "SECURITY_NAME" & DQ & ")", "")
    
    data(i, 1) = ticker 'ticker
    data(i, 2) = fbank.cells(2 + i, 5) 'isin
    
    If InStr(1, ticker, "Equity", vbTextCompare) > 1 Then
    data(i, 3) = fbank.cells(2 + i, 21) 'quantity
    Neq = Neq + 1
    Else: data(i, 3) = fbank.cells(2 + i, 10) 'quantity
    End If

    data(i, 4) = fbank.cells(2 + i, 22) ''price
    data(i, 5) = fbank.cells(2 + i, 2) ''name
    data(i, 6) = "EUR" ''Currency
    data(i, 7) = "Finint Economia Reale Italia PIR"
    data(i, 8) = "Buy"
    data(i, 9) = Format(DateToGO, "ddmmyyyy")
    data(i, 10) = fbank.cells(2 + i, 13) ''mkt value
    data(i, 11) = "IT0005261125" ''port identf
Next i
nBonds = N - Neq + 1
Workbooks(nameFiletoOpenBanca).Close SaveChanges:=False

'END working WITH banca finint excel

'open file morningstar

Close #1
    strFilePath = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Morningstar\Dati portafoglio\Template FERI - Morningstar VBA.xlsx"
    Workbooks.Open strFilePath

Set fmorn = Workbooks("Template FERI - Morningstar VBA.xlsx").Sheets("Single Line")

fmorn.cells(1, 1).Value = DateToGO
fmorn.Range(fmorn.cells(4, 2), fmorn.cells(4 + N - 1, 2)) = Application.Index(data, , 9) 'date
fmorn.Range(fmorn.cells(4, 3), fmorn.cells(4 + N - 1, 3)) = Application.Index(data, , 11) ''port identf
fmorn.Range(fmorn.cells(4, 4), fmorn.cells(4 + N - 1, 4)) = Application.Index(data, , 7)  'fund name
fmorn.Range(fmorn.cells(4, 5), fmorn.cells(4 + N - 1, 5)) = Application.Index(data, , 6) 'Currency
fmorn.Range(fmorn.cells(4, 6), fmorn.cells(4 + N - 1, 6)) = Application.Index(data, , 2) 'isin
fmorn.Range(fmorn.cells(4, 27), fmorn.cells(4 + N - 1, 27)) = Application.Index(data, , 3) 'quantity
fmorn.Range(fmorn.cells(4, 28), fmorn.cells(4 + N - 1, 28)) = Application.Index(data, , 10) 'mkt value
fmorn.Range(fmorn.cells(4, 30), fmorn.cells(4 + N - 1, 30)) = Application.Index(data, , 10) 'mkt value
fmorn.Range(fmorn.cells(4, 29), fmorn.cells(4 + N - 1, 29)) = Application.Index(data, , 6) 'Currency
fmorn.Range(fmorn.cells(4, 31), fmorn.cells(4 + N - 1, 31)) = Application.Index(data, , 8) 'Currency

Application.Calculation = xlCalculationManual
For J = 1 To N
    fmorn.cells(3 + J, 8).Value = "=BDP(""" & data(J, 1) & """,""ID_BB_GLOBAL"")"
    fmorn.cells(3 + J, 23).Value = "=BDP(""" & data(J, 1) & """,""SECURITY_NAME"")"
    fmorn.cells(3 + J, 32).Value = "=BDP(""" & data(J, 1) & """,""LONG_COMP_NAME"")"
    If J < nBonds Then
        fmorn.cells(3 + J, 22).Value = "=BDP(""" & data(J, 1) & """,""Industry_sector"")&" & """ Bond"""
        fmorn.cells(3 + J, 1).Value = "Bond"
        fmorn.cells(3 + J, 34).Value = "=BDP(""" & data(J, 1) & """,""MATURITY"")"
        fmorn.cells(3 + J, 43).Value = "=BDP(""" & data(J, 1) & """,""COUPON"")"
    Else
        fmorn.cells(3 + J, 1).Value = "=BDP(""" & data(J, 1) & """,""Security_typ2"")"
        fmorn.cells(3 + J, 22).Value = "=BDP(""" & data(J, 1) & """,""Security_typ"")"
    End If
    If InStr(1, UCase(data(J, 1)), "MTGE") > 1 Then
        fmorn.cells(3 + J, 22).Value = "=BDP(""" & data(J, 1) & """,""Security_typ"")"
        fmorn.cells(3 + J, 107).Value = "=BDP(""" & data(J, 1) & """,""MTG_FACTOR"")"
        fmorn.cells(3 + J, 112).Value = "=BDP(""" & data(J, 1) & """,""MTG_TRANCHE_TYP"")"
        fmorn.cells(3 + J, 1).Value = "=BDP(""" & data(J, 1) & """,""Security_typ2"")"
    End If
    
    If InStr(1, UCase(data(J, 1)), "EQUITY") > 1 Then
        fmorn.cells(3 + J, 30).Value = "=BDH(""" & data(J, 1) & """,""PX_LAST"",$A$1,$A$1,""Days=A,Fill =C"")"
        fmorn.cells(3 + J, 28).Value = "=AD" & J + 3 & "*AA" & J + 3
    
    End If
    
Next J

nameFiletoSaveMonrningstar = "Fondo FERI - PIR " & Format(DateToGO, "mm.yy") & "Morn VBA Formule.xlsx"
month_file1 = Format(DateToGO, "mm.yyyy")
month_filefolder = Format(DateToGO, "mm.yy")
strFolder = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Morningstar\Dati portafoglio\" & year_file & "\" & month_filefolder
createFolder (strFolder)

With Workbooks("Template FERI - Morningstar VBA.xlsx")
    .SaveAs FileName:=strFolder & "\" & nameFiletoSaveMonrningstar
    .Save
End With


Application.Calculation = xlCalculationAutomatic
End Sub
Private Function checkFolderExistance(ByVal strFolderPath As String) As Boolean

checkFolderExistance = Dir(strFolderPath, vbDirectory) <> vbNullString


End Function


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
