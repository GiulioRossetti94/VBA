Sub table_comparison()
Dim ws_d As Worksheet
Dim ws_t As Worksheet
Dim data() As Variant
Dim header(1 To 1, 1 To 7) As Variant
Dim header_month() As Variant

Dim fTable() As Variant
Dim EqtNew() As Variant
Dim BilNew() As Variant
Dim BilEqtNew() As Variant
Dim BilBondNew() As Variant
Dim FlexNew() As Variant
Dim OtheNew() As Variant
Dim IndexNew() As Variant
Dim tSort() As Variant

Application.Calculation = xlCalculationManual


Set ws_d = Sheets("Data")
Set ws_t = Sheets("Tables")
ws_t.Range("DQ4:ES100").Delete
Set dataRng = ws_d.Range(ws_d.cells(9, 397), ws_d.cells(9, 399).End(xlDown))
nFunds = ws_d.Range(ws_d.cells(9, 397), ws_d.cells(9, 397).End(xlDown)).Count

nEquity = 0
nBil = 0
nBilEqt = 0
nBilBond = 0
nFlex = 0
nOthe = 0
nIndex = 0

For i = 1 To nFunds
    If ws_d.cells(8 + i, 398) = "Azionari Italia" Then
        nEquity = nEquity + 1
    ElseIf ws_d.cells(8 + i, 398) = "Bilanciati" Then
        nBil = nBil + 1
    ElseIf ws_d.cells(8 + i, 398) = "Bilanciati Azionari" Then
        nBilEqt = nBilEqt + 1
    ElseIf ws_d.cells(8 + i, 398) = "Bilanciati Obbligazionari" Then
        nBilBond = nBilBond + 1
    ElseIf ws_d.cells(8 + i, 398) = "Flessibili" Then
        nFlex = nFlex + 1
    ElseIf ws_d.cells(8 + i, 398) = "Obbligazionari altre spec." Then
        nOthe = nOthe + 1
    ElseIf ws_d.cells(8 + i, 398) = "Indice" Then
        nIndex = nIndex + 1
        
    End If
Next i
nALL = nEquity + nBil + nBilEqt + nBilBond + nFlex + nOthe + nIndex
'Debug.Print nALL
ReDim data(1 To nFunds, 1 To 3)
data = dataRng.Value

'========================================
'INPUT BOXES
'========================================

n_month_backward = Month(Now()) - 1
startMonth = InputBox("Inserisci Data da visualizzare (in formato americano)", , Format(Now(), "mm/dd/yy"))
N_Month = InputBox("Inserisci Numero mesi", , n_month_backward)
If startMonth = vbNullString Or N_Month = vbNullString Then GoTo exitAfterCancel:
N_Month = CInt(N_Month)

'========================================
'PREPARING HEADER OF TABLE
'========================================
header(1, 1) = "ISIN"
header(1, 2) = "Ticker"
header(1, 3) = "Fondo"
header(1, 4) = "Società di gestione"
header(1, 5) = "Categoria Assogestioni"
header(1, 6) = "Totale Attivi"
header(1, 7) = "Data Avvio"

ReDim header_month(1 To 1, 1 To N_Month)
For i = 1 To N_Month
    header_month(1, i) = Format(Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(startMonth, -i) + 1, -1), "dd-mmm-yy")
    If header_month(1, i) = "30-mar-18" Then header_month(1, i) = "29-mar-18"
Next i

lastMonthOfCalc = Format(Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(CDate(header_month(1, N_Month)), -1) + 1, -1), "dd-mmm-yy")
begOfYear = Format(Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(Now(), -Month(Now())) + 1, -1), "dd-mmm-yy")
LastDayAvail = Format(Application.WorksheetFunction.WorkDay(Now(), -2))

'========================================
'PRINT HEADER OF TABLE
'========================================
ws_t.Range(ws_t.cells(4, 121), ws_t.cells(4, 127)) = header
ws_t.Range(ws_t.cells(4, 130), ws_t.cells(4, 128 + N_Month + 1)) = header_month
ws_t.cells(4, 128 + N_Month + 2) = "YTD"
ws_t.cells(4, 129) = LastDayAvail
ws_t.Range(ws_t.cells(4, 129), ws_t.cells(4, 128 + N_Month + 1)).NumberFormat = "dd-mmm-yy"
ws_t.Range(ws_t.cells(5, 129), ws_t.cells(5 + nFunds - 1, 128 + N_Month + 2)).NumberFormat = "0.00%"
ws_t.Range(ws_t.cells(5, 126), ws_t.cells(5 + nFunds - 1, 126)).NumberFormat = "_(* #,##0_)"


'========================================
'WORKING ON DATA---> PUT DATA IN ARRAY
'========================================
ReDim fTable(1 To nFunds, 1 To 8 + N_Month + 2)
For i = 1 To nFunds

    fTable(i, 1) = data(i, 1) 'isin
    If data(i, 2) = "Indice" Then
    
        fTable(i, 2) = data(i, 1) & " Index"
        fTable(i, 3) = "=BDP(""" & data(i, 1) & " Index"",""Name"")"
    '    fTable(i, 4) = "=BDP(""" & data(i, 1) & " Equity"",""fund_management_co_long"")"
        fTable(i, 4) = data(i, 3) 'name
        fTable(i, 5) = data(i, 2) 'class
        fTable(i, 6) = ""
        fTable(i, 7) = ""
        fTable(i, 9) = "=BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(130 - 1) & "4" & "," & Col_Letter(130 - 1) & "4" & ",""Days=A,Fill =C"")/" _
                        & "BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(130) & "4" & "," & Col_Letter(130) & "4" & ",""Days=A,Fill =C"")-1"
        For J = 1 To N_Month
            If J <> N_Month Then
             fTable(i, 9 + J) = "=BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(130 + J - 1) & "4" & "," & Col_Letter(130 + J - 1) & "4" & ",""Days=A,Fill =C"")/" _
                            & "BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(130 + J) & "4" & "," & Col_Letter(130 + J) & "4" & ",""Days=A,Fill =C"")-1"
             Else
             fTable(i, 9 + J) = "=BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(130 + J - 1) & "4" & "," & Col_Letter(130 + J - 1) & "4" & ",""Days=A,Fill =C"")/" _
                            & "BDH(""" & data(i, 1) & " Index"",""PX_LAST"",""" & lastMonthOfCalc & """,""" & lastMonthOfCalc & """,""Days=A,Fill =C"")-1"
            End If
        Next J
    '##############################Change here for YTD updated-->129 = last, 130= previous month####################################
        fTable(i, 9 + N_Month + 1) = "=BDH(""" & data(i, 1) & " Index"",""PX_LAST""," & Col_Letter(129) & "4" & "," & Col_Letter(129) & "4" & ",""Days=A,Fill =C"")/" _
                            & "BDH(""" & data(i, 1) & " Index"",""PX_LAST"",""" & begOfYear & """,""" & begOfYear & """,""Days=A,Fill =C"")-1"
 
    Else
    
        fTable(i, 2) = "=BDP(""" & data(i, 1) & " Equity"",""Ticker"")"
        fTable(i, 3) = "=BDP(""" & data(i, 1) & " Equity"",""Name"")"
    '    fTable(i, 4) = "=BDP(""" & data(i, 1) & " Equity"",""fund_management_co_long"")"
        fTable(i, 4) = data(i, 3) 'name
        fTable(i, 5) = data(i, 2) 'class
        fTable(i, 6) = "=BDP(""" & data(i, 1) & " Equity"",""fund_total_assets"")*1000000"
        fTable(i, 7) = "=BDP(""" & data(i, 1) & " Equity"",""fund_incept_dt"")"
    
    fTable(i, 9) = "=BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(130 - 1) & "4" & "," & Col_Letter(130 - 1) & "4" & ",""Days=A,Fill =C"")/" _
                        & "BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(130) & "4" & "," & Col_Letter(130) & "4" & ",""Days=A,Fill =C"")-1"
    For J = 1 To N_Month
        If J <> N_Month Then
         fTable(i, 9 + J) = "=BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(130 + J - 1) & "4" & "," & Col_Letter(130 + J - 1) & "4" & ",""Days=A,Fill =C"")/" _
                        & "BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(130 + J) & "4" & "," & Col_Letter(130 + J) & "4" & ",""Days=A,Fill =C"")-1"
         Else
         fTable(i, 9 + J) = "=BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(130 + J - 1) & "4" & "," & Col_Letter(130 + J - 1) & "4" & ",""Days=A,Fill =C"")/" _
                        & "BDH(""" & data(i, 1) & " Equity"",""PX_LAST"",""" & lastMonthOfCalc & """,""" & lastMonthOfCalc & """,""Days=A,Fill =C"")-1"
        End If
    Next J
'##############################Change here for YTD updated-->129 = last, 130= previous month####################################
    fTable(i, 9 + N_Month + 1) = "=BDH(""" & data(i, 1) & " Equity"",""PX_LAST""," & Col_Letter(129) & "4" & "," & Col_Letter(129) & "4" & ",""Days=A,Fill =C"")/" _
                        & "BDH(""" & data(i, 1) & " Equity"",""PX_LAST"",""" & begOfYear & """,""" & begOfYear & """,""Days=A,Fill =C"")-1"
    End If
Next i
'ws_t.Range(ws_t.Cells(5, 121), ws_t.Cells(5 + nFunds - 1, 128 + N_Month + 2)) = fTable

'========================================
' DIVIDING INTO CLASSES
'========================================
ReDim EqtNew(1 To nEquity, 1 To 8 + N_Month + 2)
ReDim BilNew(1 To nBil, 1 To 8 + N_Month + 2)
ReDim BilEqtNew(1 To nBilEqt, 1 To 8 + N_Month + 2)
ReDim BilBondNew(1 To nBilBond, 1 To 8 + N_Month + 2)
ReDim FlexNew(1 To nFlex, 1 To 8 + N_Month + 2)
ReDim OtheNew(1 To nOthe, 1 To 8 + N_Month + 2)
ReDim IndexNew(1 To nIndex, 1 To 8 + N_Month + 2)

Count = 1
For J = 1 To nFunds

    If fTable(J, 5) = "Azionari Italia" Then
'        Debug.Print j
        
        For k = 1 To 8 + N_Month + 2
            EqtNew(Count, k) = fTable(J, k)
            
        Next k
        Count = Count + 1
    End If
Next J

Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Bilanciati" Then
'        Debug.Print Count
        For k = 1 To 8 + N_Month + 2
            BilNew(Count, k) = fTable(J, k)
        Next k
    Count = Count + 1
    End If
Next J
Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Bilanciati Azionari" Then
        For k = 1 To 8 + N_Month + 2
            BilEqtNew(Count, k) = fTable(J, k)
           
        Next k
     Count = Count + 1
    End If
Next J
Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Bilanciati Obbligazionari" Then
        For k = 1 To 8 + N_Month + 2
            BilBondNew(Count, k) = fTable(J, k)
            
        Next k
    Count = Count + 1
    End If
Next J
Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Flessibili" Then
        For k = 1 To 8 + N_Month + 2
            FlexNew(Count, k) = fTable(J, k)
            
        Next k
    Count = Count + 1
    End If
Next J
Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Obbligazionari altre spec." Then
        For k = 1 To 8 + N_Month + 2
            OtheNew(Count, k) = fTable(J, k)
            
        Next k
    Count = Count + 1
    End If
Next J

Count = 1
For J = 1 To nFunds
    
    If fTable(J, 5) = "Indice" Then
        For k = 1 To 8 + N_Month + 2
            IndexNew(Count, k) = fTable(J, k)
            
        Next k
    Count = Count + 1
    End If
Next J
    

'========================================
' PRINTING DATA IN TABLE
'========================================
nSelected = 0
For cb = 7 To 13
    If ws_t.OLEObjects("CheckBox" & cb).Object.Value = True Then nSelected = nSelected + 1
Next cb

If ws_t.OLEObjects("CheckBox15").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nFunds - 1, 128 + N_Month + 2)) = fTable
End If

If nSelected = 1 Then
    If ws_t.OLEObjects("CheckBox7").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nEquity - 1, 128 + N_Month + 2)) = EqtNew
    ElseIf ws_t.OLEObjects("CheckBox8").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nBil - 1, 128 + N_Month + 2)) = BilNew
    ElseIf ws_t.OLEObjects("CheckBox9").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nBilEqt - 1, 128 + N_Month + 2)) = BilEqtNew
    ElseIf ws_t.OLEObjects("CheckBox10").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nBilBond - 1, 128 + N_Month + 2)) = BilBondNew
    ElseIf ws_t.OLEObjects("CheckBox11").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nFlex - 1, 128 + N_Month + 2)) = FlexNew
    ElseIf ws_t.OLEObjects("CheckBox12").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nOthe - 1, 128 + N_Month + 2)) = OtheNew
    ElseIf ws_t.OLEObjects("CheckBox13").Object.Value = True Then
        ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nIndex - 1, 128 + N_Month + 2)) = IndexNew
    End If
ElseIf nSelected > 1 Then
    ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nFunds - 1, 128 + N_Month + 2)) = fTable
        For cb = 7 To 13
            counter = Application.WorksheetFunction.CountA(ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5, 121).End(xlDown)))
    '        Debug.Print counter
            If ws_t.OLEObjects("CheckBox" & cb).Object.Value = False Then
                For i = counter To 1 Step -1
    '                Debug.Print ws_t.Cells(4 + counter, 125).Text; ws_t.OLEObjects("CheckBox" & cb).Object.Caption
                    If ws_t.cells(4 + i, 125).Text = ws_t.OLEObjects("CheckBox" & cb).Object.Caption Then
                         ws_t.Range(ws_t.cells(4 + i, 121), ws_t.cells(4 + i, 128 + N_Month + 2)).Delete
                    End If
                Next i
            End If
        Next cb
    End If

'CellStart = 5
'For cb = 7 To 12
'    If ws_t.OLEObjects("CheckBox" & cb).Object.Value = True Then
'        ws_t.Range(ws_t.Cells(CellStart, 121), ws_t.Cells(5 + nFunds - 1, 128 + N_Month + 2)) = fTable
'
    

'========================================
' Applying conditional formatting
'========================================
ws_t.Range(ws_t.cells(5, 121), ws_t.cells(100, 155)).FormatConditions.Delete

With ws_t.Range(ws_t.cells(5, 129), ws_t.cells(5 + nFunds - 1, 128 + N_Month + 10)).FormatConditions.Add(xlCellValue, xlLess, "=0")
    .Font.Color = rgb(192, 0, 0)
    .Font.Bold = True
End With

With ws_t.Range(ws_t.cells(5, 129), ws_t.cells(5 + nFunds - 1, 128 + N_Month + 10)).FormatConditions.Add(xlCellValue, xlGreater, "=0")
    .Font.Color = rgb(0, 176, 80)
    .Font.Bold = True
End With

With ws_t.Range(ws_t.cells(4, 121), ws_t.cells(4, 128 + N_Month + 2))
    .Interior.Color = rgb(79, 129, 189)
    .Font.Bold = True
    .Font.Color = rgb(255, 255, 255)
    .HorizontalAlignment = xlCenter
End With
rowPIR = 0
For i = 1 To nFunds
'    Debug.Print ws_t.Cells(4 + i, 122)
    If ws_t.cells(4 + i, 125).Value <> ws_t.cells(3 + i, 125).Value Then
        With ws_t.Range(ws_t.cells(4 + i, 121), ws_t.cells(4 + i, 128 + N_Month + 2)).Borders(xlEdgeTop)
            .LineStyle = xlDot
            .Weight = xlThin
        End With
    End If
    If ws_t.cells(4 + i, 122) = "FIERPIR" Then
        rowPIR = i
        With ws_t.Range(ws_t.cells(4 + i, 121), ws_t.cells(4 + i, 128 + N_Month + 2))
            .Interior.Color = rgb(220, 230, 241)
            .Font.Bold = True
        End With
    End If
Next i
'========================================
' Summary Statistics
'========================================

'ws_t.Range(ws_t.Cells(4, 130), ws_t.Cells(4, 128 + N_Month + 1)) = header_month

ws_t.Range(ws_t.cells(6, 128 + N_Month + 4), ws_t.cells(6 + N_Month - 1, 128 + N_Month + 4)) = Application.Transpose(header_month)
ws_t.Range(ws_t.cells(5, 128 + N_Month + 4), ws_t.cells(5 + N_Month, 128 + N_Month + 4)).NumberFormat = "Mmm-yy"


ws_t.Range(ws_t.cells(5, 128 + N_Month + 4), ws_t.cells(6 + N_Month, 128 + N_Month + 4)).FormatConditions.Delete
ws_t.Range(ws_t.cells(5, 128 + N_Month + 6), ws_t.cells(6 + N_Month, 128 + N_Month + 6)).FormatConditions.Delete

ws_t.cells(5, 128 + N_Month + 4) = ws_t.cells(4, 129)
ws_t.cells(6 + N_Month, 128 + N_Month + 4) = ws_t.cells(4, 128 + N_Month + 2)
ws_t.cells(4, 128 + N_Month + 5) = "Average"
ws_t.cells(4, 128 + N_Month + 6) = "Sigma"
ws_t.cells(4, 128 + N_Month + 7) = "Max"
ws_t.cells(4, 128 + N_Month + 8) = "Min"
ws_t.cells(4, 128 + N_Month + 9) = "PIR"
nFundsInTable = ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5, 121).End(xlDown)).Count

With ws_t.Range(ws_t.cells(5, 128 + N_Month + 4), ws_t.cells(6 + N_Month + 2, 128 + N_Month + 4))
    .HorizontalAlignment = xlRight
    .Font.Bold = True
    .Font.Color = rgb(0, 32, 96)
End With

With ws_t.Range(ws_t.cells(4, 128 + N_Month + 4), ws_t.cells(4, 128 + N_Month + 10))
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
    .Font.Color = rgb(0, 32, 96)
End With

For i = 1 To N_Month + 2

    ws_t.cells(4 + i, 128 + N_Month + 5).Value = "=AVERAGE(" & Col_Letter(128 + i) & "5:" & Col_Letter(128 + i) & nFundsInTable + 4 & ")"
    ws_t.cells(4 + i, 128 + N_Month + 6).Value = "=STDEV(" & Col_Letter(128 + i) & "5:" & Col_Letter(128 + i) & nFundsInTable + 4 & ")"
    ws_t.cells(4 + i, 128 + N_Month + 7).Value = "=MAX(" & Col_Letter(128 + i) & "5:" & Col_Letter(128 + i) & nFundsInTable + 4 & ")"
    ws_t.cells(4 + i, 128 + N_Month + 8).Value = "=MIN(" & Col_Letter(128 + i) & "5:" & Col_Letter(128 + i) & nFundsInTable + 4 & ")"
    ws_t.Range(ws_t.cells(4 + i, 128 + N_Month + 5), ws_t.cells(4 + i, 128 + N_Month + 9)).NumberFormat = "0.00%"
    If rowPIR <> 0 Then ws_t.cells(4 + i, 128 + N_Month + 9).Value = "=" & Col_Letter(128 + i) & rowPIR + 4
    
    On Error GoTo nextIt
    ws_t.cells(4 + i, 128 + N_Month + 4) = CDate(ws_t.cells(4 + i, 128 + N_Month + 4))
nextIt:
Next i



If ws_t.OLEObjects("CheckBox14").Object.Value = True Then

    If ws_t.OLEObjects("CheckBox15").Object.Value = True Then
        Application.OnTime Now + TimeValue("00:00:10"), "sortingByPerformance"
    Else
       Application.OnTime Now + TimeValue("00:00:4"), "sortingByPerformance"
    End If
End If
Application.Calculation = xlCalculationAutomatic

Exit Sub

exitAfterCancel:
Debug.Print "CIAO"
Application.Calculation = xlCalculationAutomatic

End Sub
Sub sortingByPerformance()
Dim ws_d As Worksheet
Dim ws_t As Worksheet
Dim tSort() As Variant

Set ws_d = Sheets("Data")
Set ws_t = Sheets("Tables")

nFundsInTable = ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5, 121).End(xlDown)).Count
N_Month = ws_t.Range(ws_t.cells(5, 130), ws_t.cells(5, 130).End(xlToRight)).Count - 1

ReDim tSort(1 To nFundsInTable, 1 To 10 + N_Month)
colSort = InputBox("Inserire numero colonna rendimenti che si vuole ordinare", "Sort", 1)
If colSort = vbNullString Then GoTo exitAfterCancel
colSort = CInt(colSort) + 8

For i = 1 To nFundsInTable
    If IsNumeric(ws_t.cells(4 + i, 121 + colSort)) = False Then
'            DoEvents
'            Application.Wait (Now + TimeValue("0:00:02"))
        Do Until Application.CalculationState = xlDone
            DoEvents
        Loop
    End If
Next i
          
Set rngFinal = ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nFundsInTable - 1, 121 + 9 + N_Month))

tSort = rngFinal.Value

QuicksortD tSort, LBound(tSort), UBound(tSort), colSort
ws_t.Range(ws_t.cells(5, 121), ws_t.cells(5 + nFundsInTable - 1, 121 + 9 + N_Month)) = tSort
rowPIR = 0
For i = 1 To nFundsInTable
    If ws_t.cells(4 + i, 122) = "FIERPIR" Then
        rowPIR = i
        With ws_t.Range(ws_t.cells(4 + i, 121), ws_t.cells(4 + i, 128 + N_Month + 2))
            .Interior.Color = rgb(220, 230, 241)
            .Font.Bold = True
        End With
        
    Else
        With ws_t.Range(ws_t.cells(4 + i, 121), ws_t.cells(4 + i, 128 + N_Month + 2))
            .Interior.Color = rgb(255, 255, 255)
            .Font.Bold = False
        End With
    End If
Next i

If rowPIR <> 0 Then
ws_t.Range(ws_t.cells(5, 128 + N_Month + 9), ws_t.cells(6 + N_Month, 128 + N_Month + 9)).Value = Application.Transpose(ws_t.Range(ws_t.cells(4 + rowPIR, 129), ws_t.cells(4 + rowPIR, 128 + N_Month + 2)))
End If

Debug.Print

Application.Calculation = xlCalculationAutomatic
Exit Sub
exitAfterCancel:

Application.Calculation = xlCalculationAutomatic
End Sub
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function
Sub QuicksortD(Ary, LB, UB, ref)
Dim M As Variant, Temp
Dim i As Long, ii As Long, iii As Integer
i = UB
ii = LB
M = Ary(Int((LB + UB) / 2), ref)
Do While ii <= i
    Do While Ary(ii, ref) > M
        ii = ii + 1
    Loop
    Do While Ary(i, ref) < M
        i = i - 1
    Loop
    If ii <= i Then
        For iii = LBound(Ary, 2) To UBound(Ary, 2)
            Temp = Ary(ii, iii): Ary(ii, iii) = Ary(i, iii)
            Ary(i, iii) = Temp
        Next
        ii = ii + 1: i = i - 1
    End If
Loop
If LB < i Then QuicksortD Ary, LB, i, ref
If ii < UB Then QuicksortD Ary, ii, UB, ref
End Sub