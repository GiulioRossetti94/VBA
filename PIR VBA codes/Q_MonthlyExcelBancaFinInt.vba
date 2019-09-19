Sub ExcelFileBancaFinInt()
Application.Calculation = xlCalculationManual
Dim ws_b As Worksheet
Dim bond_data() As Variant
Dim equity_data() As Variant
Dim etf() As Variant
Dim unlisted As New Collection

Set ws_b = ThisWorkbook.Worksheets("PTF BOND")
Set ws_p = ThisWorkbook.Worksheets("STORICO PREZZI_fi")
Set ws_e = ThisWorkbook.Worksheets("PTF EQUITY")
Set ws_f = ThisWorkbook.Worksheets("PTF ETF")

date_report = Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(Now(), -1) + 1, -1)

'WORK WITH PATRIMONIALE FERI
DateToGO = InputBox("inserisci data", , Format(date_report, "dd/mm/yyyy"))

nameFolder = Format(DateToGO, "yyyymmdd")
dateNameFile = Format(DateToGO, "dd.mm.yyyy")

year_file = Format(DateToGO, "yyyy")
month_file = Format(DateToGO, "mm yyyy")


Close #1
strFilePathPatr = "Y:\Mobiliare\08 Finint Economia Reale Italia\03_Back Office\03_Controlli NAV\" & _
                year_file & "\" & month_file & "\" & nameFolder & "\" & "PATRIMONIALE_FERI_" & dateNameFile & ".xlsx"

    Workbooks.Open strFilePathPatr
    NFile = "PATRIMONIALE_FERI_" & dateNameFile & ".xlsx"
    Set PatrFeri = Workbooks("PATRIMONIALE_FERI_" & dateNameFile & ".xlsx").Sheets("PROSPETTO QUOTA")
    Set NavCla = PatrFeri.Range("E27")
    Set NavClPir = PatrFeri.Range("E28")
    Set Liquidity = PatrFeri.Range("E16")
    Set NshareCla = PatrFeri.Range("E29")
    Set NshareClPir = PatrFeri.Range("E30")
    Liquidity = Liquidity + 0
    NshareCla = NshareCla + 0
    TotShares = NshareCla + NshareClPir
    TotNav = NavCla + NavClPir
    Workbooks(NFile).Close SaveChanges:=False


'END WORK WITH PATRIMONIALE FERI

strDate = Format(DateToGO, "yyyymmdd")

nBonds = Application.WorksheetFunction.CountA(ws_b.Range(ws_b.cells(9, 6), ws_b.cells(500, 6))) - 1
row_govi = Application.WorksheetFunction.Match("Government", ws_b.Range("B:B"), 0)
nGovi = Application.WorksheetFunction.CountIf(ws_b.Range("B:B"), "Government")
nEquity = Application.WorksheetFunction.CountA(ws_e.Range(ws_e.cells(8, 6), ws_e.cells(500, 6).End(xlDown)))
nETF = Application.WorksheetFunction.CountA(ws_f.Range(ws_f.cells(8, 12), ws_f.cells(500, 12).End(xlDown)))

ReDim equity_data(1 To nEquity, 1 To 7)
ReDim bond_data(1 To nBonds, 1 To 9)
ReDim etf(1 To nEquity, 1 To 6)

For i = 1 To nBonds - nGovi
    bond_data(i, 1) = ws_b.cells(8 + i, 6) 'ISIN
    bond_data(i, 2) = ws_b.cells(8 + i, 1) 'SECTOR
    bond_data(i, 3) = ws_b.cells(8 + i, 19) 'Nominal
    bond_data(i, 4) = ws_b.cells(8 + i, 14) 'Factor
    bond_data(i, 5) = "Titolo di debito"
    bond_data(i, 6) = ws_b.cells(8 + i, 5) 'TICKER
    
    If GetFontColor(ws_b.cells(8 + i, 5)) <> 1 Then
        bond_data(i, 7) = "S"
    Else
        bond_data(i, 7) = "N"
    End If
    
    On Error Resume Next
    Err.Clear
    bond_data(i, 7) = Application.WorksheetFunction.Index(ws_p.Range(ws_p.cells(6, 6), ws_p.cells(6, 6).End(xlDown)), _
                      Application.WorksheetFunction.Match(strDate & Trim(ws_b.cells(8 + i, 6)), ws_p.Range(ws_p.cells(6, 5), ws_p.cells(6, 5).End(xlDown)), 0))
    If Err.Number <> 0 Then
    bond_data(i, 7) = "No Price, Check Purchase date"
    End If
    
    
    If UCase(Right(ws_b.cells(8 + i, 5), 4)) = "MTGE" Then
    bond_data(i, 8) = 0
    Else: bond_data(i, 8) = 1
    End If
    'If getfontcolor(ws_b.Cells(8 + i, 1)) Then
    If ws_b.cells(8 + i, 5).Font.Color <> vbBlack Then
    bond_data(i, 9) = "N"
    Else: bond_data(i, 9) = "S"
    End If
Next i

For i = 1 To nEquity
    equity_data(i, 1) = ws_e.cells(7 + i, 5) 'ISIN
    equity_data(i, 2) = ws_e.cells(7 + i, 1) 'SECTOR
    equity_data(i, 3) = ws_e.cells(7 + i, 7) 'Quantity
    equity_data(i, 4) = "Titolo di capitale"
    equity_data(i, 5) = ws_e.cells(7 + i, 4) 'TICKER
    
    equity_data(i, 6) = "S"

'    End If
Next i

For i = 1 To nETF
    etf(i, 1) = ws_f.cells(7 + i, 5) 'ISIN
    etf(i, 2) = ws_f.cells(7 + i, 1) 'SECTOR
    etf(i, 3) = ws_f.cells(7 + i, 7) 'Quantity
    etf(i, 4) = "Titolo di capitale"
    etf(i, 5) = ws_f.cells(7 + i, 4) 'TICKER
    etf(i, 6) = "S"
Next i




indx = 0
For i = nBonds - nGovi To nBonds - 1
     
    bond_data(i, 1) = ws_b.cells(row_govi + indx, 6) 'ISIN
    bond_data(i, 2) = ws_b.cells(row_govi + indx, 1) 'SECTOR
    bond_data(i, 3) = ws_b.cells(row_govi + indx, 19) 'Nominal
    bond_data(i, 4) = ws_b.cells(row_govi + indx, 14) 'Factor
    
    bond_data(i, 5) = "Titolo di debito"
    bond_data(i, 6) = ws_b.cells(row_govi + indx, 5) 'TICKER
    On Error Resume Next
    Err.Clear
    bond_data(i, 7) = Application.WorksheetFunction.Index(ws_p.Range(ws_p.cells(row_govi, 6), ws_p.cells(row_govi, 6).End(xlDown)), _
                      Application.WorksheetFunction.Match(strDate & Trim(ws_b.cells(row_govi + indx, 6)), ws_p.Range(ws_p.cells(row_govi, 5), ws_p.cells(row_govi, 5).End(xlDown)), 0))
    If Err.Number <> 0 Then
    bond_data(i, 7) = "No Price, Check Purchase date"
    End If
    indx = indx + 1
'    Debug.Print bond_data(i, 1), bond_data(i, 7), bond_data(i, 2)
     bond_data(i, 8) = 1
     
Next i


Close #1
    strFilePath = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Banca Finint\Dati portafoglio\Template FERI - Banca Finint VBA.xlsx"
    Workbooks.Open strFilePath

Set Outp = Workbooks("Template FERI - Banca Finint VBA.xlsx").Sheets("Composizione PTF Fondo")
Set OutpPage1 = Workbooks("Template FERI - Banca Finint VBA.xlsx").Sheets("Partecipanti Gruppo")

Outp.cells(1, 1) = "Composizione Portafoglio del Fondo al " & DateToGO
Outp.cells(1, 5) = DateToGO
Outp.Range(Outp.cells(3, 5), Outp.cells(2 + nBonds - 1, 5)) = Application.Index(bond_data, , 1)
Outp.Range(Outp.cells(3, 7), Outp.cells(2 + nBonds - 1, 7)) = Application.Index(bond_data, , 2)
Outp.Range(Outp.cells(3, 10), Outp.cells(2 + nBonds - 1, 10)) = Application.Index(bond_data, , 3)
Outp.Range(Outp.cells(3, 24), Outp.cells(2 + nBonds - 1, 24)) = Application.Index(bond_data, , 4)
Outp.Range(Outp.cells(3, 1), Outp.cells(2 + nBonds - 1, 1)) = Application.Index(bond_data, , 5)
Outp.Range(Outp.cells(3, 22), Outp.cells(2 + nBonds - 1, 22)) = Application.Index(bond_data, , 7)
Outp.Range(Outp.cells(3, 22), Outp.cells(2 + nBonds - 1, 22)) = Application.Index(bond_data, , 7)
Outp.Range(Outp.cells(3, 8), Outp.cells(2 + nBonds - 1, 8)) = Application.Index(bond_data, , 9)
Outp.Range(Outp.cells(3, 38), Outp.cells(2 + nBonds - 1, 38)) = Application.Index(bond_data, , 6)

'Equity
Outp.Range(Outp.cells(3 + nBonds - 1, 5), Outp.cells(3 + nBonds + nEquity - 2, 5)) = Application.Index(equity_data, , 1)
Outp.Range(Outp.cells(3 + nBonds - 1, 7), Outp.cells(3 + nBonds + nEquity - 2, 7)) = Application.Index(equity_data, , 2)
Outp.Range(Outp.cells(3 + nBonds - 1, 1), Outp.cells(3 + nBonds + nEquity - 2, 1)) = Application.Index(equity_data, , 4)
Outp.Range(Outp.cells(3 + nBonds - 1, 8), Outp.cells(3 + nBonds + nEquity - 2, 8)) = Application.Index(equity_data, , 6)
Outp.Range(Outp.cells(3 + nBonds - 1, 21), Outp.cells(3 + nBonds + nEquity - 2, 21)) = Application.Index(equity_data, , 3)
Outp.Range(Outp.cells(3 + nBonds - 1, 38), Outp.cells(3 + nBonds + nEquity - 2, 38)) = Application.Index(equity_data, , 5)
Outp.Range(Outp.cells(3 + nBonds - 1, 1), Outp.cells(3 + nBonds + nEquity - 2, 22)).Interior.Color = rgb(218, 238, 243)
Debug.Print nBonds
'etf
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 5), Outp.cells(3 + nBonds + nEquity + nETF - 2, 5)) = Application.Index(etf, , 1)
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 1), Outp.cells(3 + nBonds + nEquity + nETF - 2, 1)) = Application.Index(etf, , 4)
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 8), Outp.cells(3 + nBonds + nEquity + nETF - 2, 8)) = Application.Index(etf, , 6)
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 21), Outp.cells(3 + nBonds + nEquity + nETF - 2, 21)) = Application.Index(etf, , 3)
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 38), Outp.cells(3 + nBonds + nEquity + nETF - 2, 38)) = Application.Index(etf, , 5)
Outp.Range(Outp.cells(3 + nBonds + nEquity - 1, 1), Outp.cells(3 + nBonds + nEquity + nETF - 2, 22)).Interior.Color = rgb(184, 204, 228)



With Outp.Range(Outp.cells(3 + nBonds - 1, 21), Outp.cells(3 + nBonds + nEquity + nETF - 2, 22)).Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With

For J = 1 To nBonds - 1

    Outp.cells(2 + J, 2).Value = "=BDP(""" & bond_data(J, 6) & """,""SECURITY_NAME"")"
    Outp.cells(2 + J, 4).Value = "=BDP(""" & bond_data(J, 6) & """,""COMPANY_TAX_IDENTIFIER"")"
    Outp.cells(2 + J, 6).Value = "=BDP(""" & bond_data(J, 6) & """,""COUNTRY_FULL_NAME"")"
    Outp.cells(2 + J, 7).Value = "=BDP(""" & bond_data(J, 6) & """,""INDUSTRY_SECTOR"")"
    Outp.cells(2 + J, 26).Value = "=BDP(""" & bond_data(J, 6) & """,""IS_PERPETUAL"")"
    Outp.cells(2 + J, 27).Value = "=BDP(""" & bond_data(J, 6) & """,""BULLET"")"
    Outp.cells(2 + J, 28).Value = "=BDP(""" & bond_data(J, 6) & """,""SINKABLE"")"
    Outp.cells(2 + J, 29).Value = "=BDP(""" & bond_data(J, 6) & """,""CALLABLE"")"
    Outp.cells(2 + J, 30).Value = "=BDP(""" & bond_data(J, 6) & """,""CPN_TYP"")"
    'Outp.Cells(2 + j, 32).Value = "=IF(AD" & i + 2 & "=""FIXED"",""-"",BDP(" & bond_data(j, 6) & """,""CPN_FREQ"")"
    Outp.cells(2 + J, 32).Value = "=IF(AD" & J + 2 & "=""FIXED"",""-"",BDP(""" & bond_data(J, 6) & """,""CPN_FREQ""))"
    Outp.cells(2 + J, 33).Value = "=IF(AD" & J + 2 & "=""FIXED"",""-"",BDP(""" & bond_data(J, 6) & """,""REFIX_FREQ""))"
    Outp.cells(2 + J, 34).Value = "=BDP(""" & bond_data(J, 6) & """,""CALC_TYP_DES"")"
    Outp.cells(2 + J, 18).Value = "=BDP(""" & bond_data(J, 6) & """,""DUR_ADJ_MID"")"
    Outp.cells(2 + J, 19).Value = "=IF(ISNUMBER(R" & J + 2 & "),"""",BDP(""" & bond_data(J, 6) & """,""MTG_WAL""))"
    Outp.cells(2 + J, 13).Value = "=IFERROR(J" & J + 2 & "*(1/100)*V" & J + 2 & "*X" & J + 2 & ","""")"
    Outp.cells(2 + J, 16).Formula = "=AD" & J + 2
'    Outp.Cells(2 + j, 15).Formula = "=IF(AC" & j + 2 & "=""Y"",IF(AB" & j + 2 & "=""Y"",""Sinkable"",IF(AA" & j + 2 & "=""Y"",""Bullet"",""Perpetual""))))"
    Outp.cells(2 + J, 15).Formula = "=IF(AC" & J + 2 & "=""Y"",IF(AB" & J + 2 & "=""Y"",""Sinkable"",IF(AA" & J + 2 & "=""Y"",""Bullet"",""Perpetual"")),"""")"


Next J


For J = 1 To nEquity

    Outp.cells(3 + nBonds - 2 + J, 2).Value = "=BDP(""" & equity_data(J, 5) & """,""NAME"")"
    Outp.cells(3 + nBonds - 2 + J, 22).Value = "=BDH(""" & equity_data(J, 5) & """,""PX_LAST"",E1,E1,""Days=A,Fill =C"")"
    Outp.cells(3 + nBonds - 2 + J, 6).Value = "=BDP(""" & equity_data(J, 5) & """,""COUNTRY_FULL_NAME"")"
    Outp.cells(3 + nBonds - 2 + J, 11).Value = "=IF(ISNUMBER(BDP(""" & equity_data(J, 5) & """,""TOTAL_EQUITY"")),BDP(""" & equity_data(J, 5) & """,""TOTAL_EQUITY""),"""")"
    Outp.cells(3 + nBonds - 2 + J, 12).Value = "=IFERROR(J" & 3 + nBonds - 2 + J & "/K" & 3 + nBonds - 2 + J & ","""")"
    Outp.cells(3 + nBonds - 2 + J, 13).Value = "=IFERROR(U" & 3 + nBonds - 2 + J & "*V" & 3 + nBonds - 2 + J & ","""")"
    Outp.cells(3 + nBonds - 2 + J, 10).Value = "=IFERROR(U" & 3 + nBonds - 2 + J & "*(1/1000000)*V" & 3 + nBonds - 2 + J & ","""")"
    
Next J

For J = 1 To nETF

    Outp.cells(3 + nBonds - 2 + nEquity + J, 2).Value = "=BDP(""" & etf(J, 5) & """,""NAME"")"
    Outp.cells(3 + nBonds - 2 + nEquity + J, 22).Value = "=BDH(""" & etf(J, 5) & """,""PX_LAST"",$E$1,$E$1,""Days=A,Fill =C"")"
    Outp.cells(3 + nBonds - 2 + nEquity + J, 6).Value = "=BDP(""" & etf(J, 5) & """,""COUNTRY_FULL_NAME"")"
    Outp.cells(3 + nBonds - 2 + nEquity + J, 13).Value = "=IFERROR(U" & 3 + nBonds + nEquity - 2 + J & "*V" & 3 + nBonds + nEquity - 2 + J & ","""")"
    Outp.cells(3 + nBonds - 2 + nEquity + J, 10).Value = "=IFERROR(U" & 3 + nBonds + nEquity - 2 + J & "*(1/1000000)*V" & 3 + nBonds + nEquity - 2 + J & ","""")"
    
Next J


'Liquidity
Outp.Range(Outp.cells(3 + nBonds + nEquity + nETF - 1, 1), Outp.cells(3 + nBonds + nEquity + nETF, 22)).Interior.Color = rgb(252, 213, 180)
Outp.Range(Outp.cells(3 + nBonds + nEquity + nETF - 1, 1), Outp.cells(3 + nBonds + nEquity + nETF, 1)) = "Liquidità"

Outp.cells(3 + nBonds + nEquity + nETF - 1, 10).Value = "=U" & 3 + nBonds + nEquity + nETF & "-J" & 3 + nBonds + nEquity + nETF
Outp.cells(3 + nBonds + nEquity + nETF - 1, 13).Value = "=j" & 3 + nBonds + nEquity + nETF - 1
Outp.cells(3 + nBonds + nEquity + nETF, 13).Value = "=j" & 3 + nBonds + nEquity + nETF

Outp.cells(3 + nBonds + nEquity + nETF - 1, 2) = "Liquidità State Street"
Outp.cells(3 + nBonds + nEquity + nETF, 2) = "Liquidità Banca FinInt"
Outp.cells(3 + nBonds + nEquity + nETF, 21).Value = Liquidity

OutpPage1.cells(3, 3) = TotShares
OutpPage1.cells(3, 5) = TotNav
OutpPage1.cells(3, 2) = NshareCla
OutpPage1.cells(1, 1) = "Situazione al " & DateToGO



nameFiletoSave = "Fondo FERI - PIR " & Format(DateToGO, "mm.yy") & " VBA Formule.xlsx"
month_file1 = Format(DateToGO, "mm.yyyy")
month_filefolder = Format(DateToGO, "mm.yy")
strFolder = "Y:\Mobiliare\08 Finint Economia Reale Italia\02_Middle Office\Banca Finint\Dati portafoglio\" & year_file & "\" & month_filefolder
createFolder (strFolder)

With Workbooks("Template FERI - Banca Finint VBA.xlsx")
    .SaveAs FileName:=strFolder & "\" & nameFiletoSave
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

Function GetFontColor(ByVal Target As Range) As Integer
    GetFontColor = Target.Font.ColorIndex
End Function