Sub bloombergExcel()
Dim fbank As Worksheet
Dim data() As Variant
Dim dataToPaste() As Variant

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
Liquidity = fbank.cells(2 + N + 2, 21)
Nav = Workbooks(nameFiletoOpenBanca).Sheets("Partecipanti Gruppo").Range("E3")

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
'working with bloomberg file

Close #1
    strFilePath = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Bloomberg\Dati portafoglio\Template FERI - Bloomberg VBA.xls"
    Workbooks.Open strFilePath
    
    Set fbloom = Workbooks("Template FERI - Bloomberg VBA.xls").Sheets("VBA BBG")

'fbloom.Range("C3") = Format(date_report, "dd/mm/yyyy")
'fbloom.Range("E3") = Nav
'fbloom.Range("F3") = Liquidity

fbloom.Range(fbloom.cells(6, 1), fbloom.cells(6 + N - 1, 1)) = "FIERITA IM Equity"
fbloom.Range(fbloom.cells(6, 2), fbloom.cells(6 + N - 1, 2)) = "Finint Economia Reale Italia - Classe A"
fbloom.Range(fbloom.cells(6, 3), fbloom.cells(6 + N - 1, 3)) = Format(date_report, "dd/mm/yyyy")
fbloom.Range(fbloom.cells(6, 4), fbloom.cells(6 + N - 1, 4)) = Application.Index(data, , 6) 'currency
fbloom.Range(fbloom.cells(6, 5), fbloom.cells(6 + N - 1, 5)) = Nav
fbloom.Range(fbloom.cells(6, 6), fbloom.cells(6 + N - 1, 6)) = Liquidity
fbloom.Range(fbloom.cells(6, 8), fbloom.cells(6 + N - 1, 8)) = Application.Index(data, , 2) 'isin
fbloom.Range(fbloom.cells(6, 9), fbloom.cells(6 + N - 1, 9)) = Application.Index(data, , 3) 'quantity
fbloom.Range(fbloom.cells(6, 10), fbloom.cells(6 + N - 1, 10)) = Application.Index(data, , 10) 'mkt value
fbloom.Range(fbloom.cells(6, 15), fbloom.cells(6 + N - 1, 15)) = Application.Index(data, , 1) 'ticker
fbloom.Range(fbloom.cells(6, 16), fbloom.cells(6 + N - 1, 16)) = Application.Index(data, , 4) 'price

Application.Calculation = xlCalculationManual
For J = 1 To N
    fbloom.cells(5 + J, 7).Value = "=BDP(""" & data(J, 1) & """,""SECURITY_NAME"")"

    If J < nBonds Then
        fbloom.cells(5 + J, 12).Value = "=BDP(""" & data(J, 1) & """,""MATURITY"")"
        fbloom.cells(5 + J, 13).Value = "=BDP(""" & data(J, 1) & """,""COUPON"")"
    End If
    If InStr(1, UCase(data(J, 1)), "MTGE") > 1 Then
        fbloom.cells(5 + J, 17).Value = "=BDP(""" & data(J, 1) & """,""MTG_FACTOR"")"
    End If
    If InStr(1, UCase(data(J, 1)), "EQUITY") > 1 Then
        fbloom.cells(5 + J, 16).Value = "=BDH(""" & data(J, 1) & """,""PX_LAST"",$C$6,$C$6,""Days=A,Fill =C"")"
    End If
    
    If fbloom.cells(5 + J, 10) = "" Then
        fbloom.cells(5 + J, 10).Value = "=I" & J + 5 & "*P" & J + 5
    End If


    
    fbloom.cells(5 + J, 11).Value = "=J" & J + 5 & "/E" & J + 5
    
Next J

Set RangeToCopy = fbloom.Range(fbloom.cells(6, 3), fbloom.cells(6 + N - 1, 17))
ReDim dataToPaste(1 To N, 1 To 15)
dataToPaste = RangeToCopy.Value

fbloom.Range(fbloom.cells(6 + N, 3), fbloom.cells(6 + N + N - 1, 17)) = dataToPaste
fbloom.Range(fbloom.cells(6 + N, 1), fbloom.cells(6 + N + N - 1, 1)) = "FIERPIR IM Equity"
fbloom.Range(fbloom.cells(6 + N, 2), fbloom.cells(6 + N + N - 1, 2)) = "Finint Economia Reale Italia - Classe PIR"

For J = 1 To N
    fbloom.cells(5 + J + N, 7).Value = "=BDP(""" & data(J, 1) & """,""SECURITY_NAME"")"

    If J < nBonds Then
        fbloom.cells(5 + J + N, 12).Value = "=BDP(""" & data(J, 1) & """,""MATURITY"")"
        fbloom.cells(5 + J + N, 13).Value = "=BDP(""" & data(J, 1) & """,""COUPON"")"
    End If
    If InStr(1, UCase(data(J, 1)), "MTGE") > 1 Then
        fbloom.cells(5 + J + N, 17).Value = "=BDP(""" & data(J, 1) & """,""MTG_FACTOR"")"
    End If
    
    If InStr(1, UCase(data(J, 1)), "EQUITY") > 1 Then
        fbloom.cells(5 + J + N, 16).Value = "=BDH(""" & data(J, 1) & """,""PX_LAST"",$C$6,$C$6,""Days=A,Fill =C"")"
    End If
    
    If IsEmpty(fbloom.cells(5 + J + N, 10)) Then
        fbloom.cells(5 + J + N, 10).Value = "=I" & J + 5 + N & "*P" & J + 5 + N
    End If


    fbloom.cells(5 + J + N, 11).Value = "=J" & J + 5 + N & "/E" & J + 5 + N
    
Next J

nameFiletoSaveBloomberg = "Fondo FERI - PIR " & Format(DateToGO, "mm.yy") & "BBG VBA Formule.xlsx"
month_file1 = Format(DateToGO, "mm.yyyy")
month_filefolder = Format(DateToGO, "mm.yy")
strFolder = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Bloomberg\Dati portafoglio\" & year_file & "\" & month_filefolder
createFolder (strFolder)

With Workbooks("Template FERI - Bloomberg VBA.xls")
    .SaveAs FileName:=strFolder & "\" & nameFiletoSaveBloomberg
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

