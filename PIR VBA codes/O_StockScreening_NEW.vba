Option Base 0
Sub Screening_sorting()

Call DeleteCells

Dim StartTime As Double
Dim SecondsElapsed As Double

Dim fact() As Variant
Dim isis_mat() As Variant
Dim fact_col() As Variant
Dim ut_array() As Variant
Dim name_mat() As Variant
Dim ticker_mat() As Variant
Dim ISIN_sorted() As Variant
Dim ind_mat() As Variant
Dim index_mat() As Variant

Dim w_mn As Worksheet
Dim w_ns As Worksheet
Dim w_c As Worksheet
Dim key() As Variant
Dim All_ones_or_zeros() As Variant
Dim sumScore() As Variant
Dim rankSum() As Variant

StartTime = Timer

Name_to_check = "log"
flag = False
For Each Sheet In Worksheets
    If Name_to_check = Sheet.Name Then
        flag = True
    End If
Next Sheet
If flag = False Then
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count - 1))
    ws.Name = "log"
End If

Set w_mn = Sheets("Monitor Azioni")
Set w_ns = Sheets("log")
Set w_c = Sheets("Tab")
Set ISIN = w_mn.Range(w_mn.cells(7, 2), w_mn.cells(7, 2).End(xlDown))
Set ticker = w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown))
Set Names_stocks = w_mn.Range(w_mn.cells(7, 4), w_mn.cells(7, 4).End(xlDown))
Set indName = w_mn.Range(w_mn.cells(7, 7), w_mn.cells(7, 7).End(xlDown))
Set indexName = w_mn.Range(w_mn.cells(7, 9), w_mn.cells(7, 9).End(xlDown))

Call DeleteCellsLOG

ticker_mat = ticker.Value
isin_mat = ISIN.Value
name_mat = Names_stocks.Value
ind_mat = indName.Value
index_mat = indexName.Value

nColDeltaTarget = 25

N = ISIN.Rows.Count         'number of stocks
'Debug.Print n
nfact = w_mn.Range("I1")    'number of factors

   
'Collect column indices of the factors of interest
ReDim fact_col(1 To nfact)

J = 1
For i = 10 To 200
    If w_mn.cells(1, i) = 1 Then
        fact_col(J) = w_mn.cells(1, i).Column
'        Debug.Print fact_col(j)
        J = J + 1
        
    End If
Next i

ReDim fact(1 To N, 1 To nfact)

For i = LBound(fact, 2) To UBound(fact, 2)
    For J = LBound(fact, 1) To UBound(fact, 1)

        fact(J, i) = w_mn.cells(6 + J, fact_col(i))
    Next J
Next i
'Debug.Print UBound(fact, 2)



'Debug.Print UBound(isin_mat, 1);
'Debug.Print LBound(isin_mat, 2)


ReDim ut_array(1 To N, 1 To 2)
ReDim ISIN_sorted(1 To 60, 1 To nfact)


'sorting and selecting top20 isin
For J = LBound(fact, 2) To UBound(fact, 2)
    For i = LBound(ut_array, 1) To UBound(ut_array, 1)
        If IsNumeric(fact(i, J)) Then
            ut_array(i, 1) = fact(i, J)
        Else
            ut_array(i, 1) = -9999
        End If
        ut_array(i, 2) = isin_mat(i, 1)
    Next i
    QuicksortD ut_array, LBound(ut_array), UBound(ut_array), 1
    For kappa = 1 To 60
        ISIN_sorted(kappa, J) = ut_array(kappa, 2)
    Next kappa
    
Next J

ReDim key(1 To nfact)
For i = 1 To nfact
    If w_mn.cells(5, fact_col(i)).MergeCells Then
        a = Split(w_mn.cells(5, fact_col(i)).MergeArea.Address, ":")
        ass = Trim(w_mn.Range(a(0)).Text)
    Else
        ass = Trim(w_mn.cells(6, fact_col(i)).Text)
    End If
       
    key(i) = ass
   
Next i

'***********************
'copy and paste formats and stocks (first 9 columns
'***********************

'w_mn.Range(w_mn.Cells(6, 2), w_mn.Cells(6 + n, 9)).Copy
'w_c.Cells(3, 1).PasteSpecial xlPasteValuesAndNumberFormats
'w_c.Cells(3, 1).PasteSpecial xlPasteFormats
Dim first9col As Variant
ReDim first9col(1 To N, 1 To 9)
Set description_stock = w_mn.Range(w_mn.cells(6, 2), w_mn.cells(6 + N, 9))
first9col = description_stock.Value
w_c.Range(w_c.cells(3, 1), w_c.cells(3 + N, 9)) = first9col

w_c.Range("A:H").Columns.ColumnWidth = 18
w_c.Range(w_c.cells(3, 9), w_c.cells(3, 8 + nfact)) = key
w_c.Range(w_c.Columns(9), w_c.Columns(9 + nfact)).Columns.ColumnWidth = 13
Application.CutCopyMode = False


'***********************
'zeros and ones table
'***********************
ReDim All_ones_or_zeros(1 To N, 1 To nfact)
For J = 1 To nfact
    On Error Resume Next
    Err.Clear
    For i = 1 To 60
        row_index = Application.WorksheetFunction.Match(ISIN_sorted(i, J), w_c.Range(w_c.cells(4, 1), w_c.cells(3 + N, 1)), 0)
        If Err.Number = 0 Then
'            w_c.Cells(row_index + 3, 8 + j) = 1
'            w_c.Cells(row_index + 3, 8 + j).NumberFormat = "0"
             All_ones_or_zeros(row_index, J) = 1
        Else: GoTo nit
        End If
nit:
    Next i

Next J


w_c.Range(w_c.cells(4, 9), w_c.cells(3 + N, 8 + nfact)) = All_ones_or_zeros

sumCol = 9 + nfact + 1
rankCol = sumCol + 3

w_c.cells(3, sumCol) = "SUM"
w_c.cells(3, rankCol) = "SUM"
w_c.cells(3, rankCol + 1) = "TICKER"
w_c.cells(3, rankCol + 2) = "NAME"
w_c.cells(3, rankCol + 3) = "INDUSTRY"
w_c.cells(3, rankCol + 4) = "INDEX"

ReDim sumScore(1 To N, 1 To 1)

For i = LBound(All_ones_or_zeros, 1) To UBound(All_ones_or_zeros, 1)
    rowSum = 0
    For J = LBound(All_ones_or_zeros, 2) To UBound(All_ones_or_zeros, 2)
        rowSum = rowSum + All_ones_or_zeros(i, J)
    Next J
    sumScore(i, 1) = rowSum
'    Debug.Print sumScore(i, 1)
Next i
w_c.Range(w_c.cells(4, sumCol), w_c.cells(3 + N, sumCol)) = sumScore

'rank sum

ReDim rankSum(1 To N, 1 To 5)
For i = LBound(rankSum, 1) To UBound(rankSum, 1)
    rankSum(i, 1) = sumScore(i, 1)
    rankSum(i, 3) = name_mat(i, 1)
    rankSum(i, 2) = ticker_mat(i, 1)
    rankSum(i, 4) = ind_mat(i, 1)
    rankSum(i, 5) = index_mat(i, 1)
Next i

QuicksortD rankSum, LBound(rankSum), UBound(rankSum), 1

rankCol = sumCol + 3

'Dim stockInPort() As Variant
'Set stockArray = Sheets("PTF EQUITY").Range(Sheets("PTF EQUITY").Cells(8, 4), Sheets("PTF EQUITY").Cells(8, 4).End(xlDown))
'nEQT1 = Sheets("PTF EQUITY").Range(Sheets("PTF EQUITY").Cells(8, 4), Sheets("PTF EQUITY").Cells(8, 4).End(xlDown)).Count
'
'ReDim stockInPort(1 To nEQT1, 1 To 1)
'stockInPort = stockArray.Value
'
'
'For i = 1 To n
'    For j = 1 To nEQT1
'        If stockInPort(j, 1) = rankSum(i, 2) Then
'        rankSum(i, 2).Font.Color = RGB(0, 255, 0)
'        End If
'    Next j
'Next i


w_c.Range(w_c.cells(4, rankCol), w_c.cells(3 + N, rankCol + 4)) = rankSum

Dim matSubInd(1 To 23, 1 To 60) As Variant

A1 = 1
a2 = 1
a3 = 1
a4 = 1
a5 = 1
a6 = 1
a7 = 1
a8 = 1
a9 = 1
a10 = 1
For i = 1 To 60
    If rankSum(i, 4) = "Consumer, Non-cyclical" Then
        matSubInd(A1, 1) = rankSum(i, 1)
        matSubInd(A1, 2) = rankSum(i, 2)
        matSubInd(A1, 3) = rankSum(i, 3)
        matSubInd(A1, 4) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(A1, 2), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(A1, 5) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(A1, 2), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(A1, 6) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(A1, 2), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
                    
        'matSubInd(a1, 5)
        
        A1 = A1 + 1
    ElseIf rankSum(i, 4) = "Energy" Then
        matSubInd(a2, 7) = rankSum(i, 1)
        matSubInd(a2, 8) = rankSum(i, 2)
        matSubInd(a2, 9) = rankSum(i, 3)
        matSubInd(a2, 10) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a2, 8), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a2, 11) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a2, 8), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a2, 12) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a2, 8), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))

        a2 = a2 + 1
    ElseIf rankSum(i, 4) = "Industrial" Then
        matSubInd(a3, 13) = rankSum(i, 1)
        matSubInd(a3, 14) = rankSum(i, 2)
        matSubInd(a3, 15) = rankSum(i, 3)
        matSubInd(a3, 16) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a3, 14), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a3, 17) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a3, 14), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a3, 18) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a3, 14), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a3 = a3 + 1
    ElseIf rankSum(i, 4) = "Consumer, Cyclical" Then
        matSubInd(a4, 19) = rankSum(i, 1)
        matSubInd(a4, 20) = rankSum(i, 2)
        matSubInd(a4, 21) = rankSum(i, 3)
        matSubInd(a4, 22) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a4, 20), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a4, 23) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a4, 20), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a4, 24) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a4, 20), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a4 = a4 + 1
    ElseIf rankSum(i, 4) = "Communications" Then
        matSubInd(a5, 25) = rankSum(i, 1)
        matSubInd(a5, 26) = rankSum(i, 2)
        matSubInd(a5, 27) = rankSum(i, 3)
        matSubInd(a5, 28) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a5, 26), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a5, 29) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a5, 26), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a5, 30) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a5, 26), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a5 = a5 + 1
    ElseIf rankSum(i, 4) = "Utilities" Then
        matSubInd(a6, 31) = rankSum(i, 1)
        matSubInd(a6, 32) = rankSum(i, 2)
        matSubInd(a6, 33) = rankSum(i, 3)
        matSubInd(a6, 34) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a6, 32), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a6, 35) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a6, 32), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a6, 36) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a6, 32), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a6 = a6 + 1
    ElseIf rankSum(i, 4) = "Financial" Then
        matSubInd(a7, 37) = rankSum(i, 1)
        matSubInd(a7, 38) = rankSum(i, 2)
        matSubInd(a7, 39) = rankSum(i, 3)
        matSubInd(a7, 40) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a7, 38), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a7, 41) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a7, 38), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a7, 42) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a7, 38), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))

        a7 = a7 + 1
    ElseIf rankSum(i, 4) = "Technology" Then
        matSubInd(a8, 43) = rankSum(i, 1)
        matSubInd(a8, 44) = rankSum(i, 2)
        matSubInd(a8, 45) = rankSum(i, 3)
        matSubInd(a8, 46) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a8, 44), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a8, 47) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a8, 44), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a8, 48) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a8, 44), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a8 = a8 + 1
    ElseIf rankSum(i, 4) = "Basic Materials" Then
        matSubInd(a9, 49) = rankSum(i, 1)
        matSubInd(a9, 50) = rankSum(i, 2)
        matSubInd(a9, 51) = rankSum(i, 3)
        matSubInd(a9, 52) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a9, 50), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a9, 53) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a9, 50), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a9, 54) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a9, 50), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))

        a9 = a9 + 1
    Else: rankSum(i, 4) = "Diversified"
        matSubInd(a10, 55) = rankSum(i, 1)
        matSubInd(a10, 56) = rankSum(i, 2)
        matSubInd(a10, 57) = rankSum(i, 3)
        matSubInd(a10, 58) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 10), w_mn.cells(7, 10).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a10, 56), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0)) / 1000000
        matSubInd(a10, 59) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, nColDeltaTarget), w_mn.cells(7, nColDeltaTarget).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a10, 56), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        matSubInd(a10, 60) = Application.WorksheetFunction.Index(w_mn.Range(w_mn.cells(7, 12), w_mn.cells(7, 12).End(xlDown)), _
                            Application.WorksheetFunction.Match(matSubInd(a10, 56), w_mn.Range(w_mn.cells(7, 3), w_mn.cells(7, 3).End(xlDown)), 0))
        
        a10 = a10 + 1
    End If
Next i


Dim matNameIndText(1 To 1, 1 To 10)
matNameIndText(1, 1) = "Consumer, Non-cyclical"
matNameIndText(1, 2) = "Energy"
matNameIndText(1, 3) = "Industrial"
matNameIndText(1, 4) = "Consumer, Cyclical"
matNameIndText(1, 5) = "Communications"
matNameIndText(1, 6) = "Utilities"
matNameIndText(1, 7) = "Financial"
matNameIndText(1, 8) = "Technology"
matNameIndText(1, 9) = "Basic Materials"
matNameIndText(1, 10) = "Diversified"

Dim headTitles(1 To 1, 1 To 6) As Variant
headTitles(1, 1) = "Score"
headTitles(1, 2) = "Ticker"
headTitles(1, 3) = "Name"
headTitles(1, 4) = "Cap (in Millions)"
headTitles(1, 5) = "Pot"
headTitles(1, 6) = "Perf (YTD)"

nColSubInd = rankCol + 6
h = 0
arrSubInd = 1

While h < 60
   
   w_c.cells(3, nColSubInd + h + 2).Value = matNameIndText(1, arrSubInd)
   w_c.Range(w_c.cells(4, nColSubInd + h), w_c.cells(4, nColSubInd + h + 5)) = headTitles
   w_c.Range(w_c.cells(3, nColSubInd + h), w_c.cells(23, nColSubInd + h)).Borders(xlEdgeLeft).LineStyle = xlContinuous
   w_c.Range(w_c.cells(3, nColSubInd + h), w_c.cells(23, nColSubInd + h)).HorizontalAlignment = xlCenter
   w_c.Range(w_c.cells(4, nColSubInd + h), w_c.cells(23, nColSubInd + h)).NumberFormat = "#"
   w_c.Range(w_c.cells(4, nColSubInd + h + 3), w_c.cells(23, nColSubInd + h + 3)).NumberFormat = "#,##0"
   w_c.Range(w_c.cells(4, nColSubInd + h + 4), w_c.cells(23, nColSubInd + h + 4)).NumberFormat = "#.00%"
   w_c.Range(w_c.cells(4, nColSubInd + h + 4), w_c.cells(4, nColSubInd + h + 5)).HorizontalAlignment = xlCenter
   w_c.Range(w_c.cells(4, nColSubInd + h + 5), w_c.cells(23, nColSubInd + h + 5)).NumberFormat = "#.00%"
   w_c.Range(w_c.cells(5, nColSubInd + h + 4), w_c.cells(23, nColSubInd + h + 4)).HorizontalAlignment = xlRight
   arrSubInd = arrSubInd + 1
   h = h + 6
Wend
w_c.Range(w_c.cells(5, nColSubInd), w_c.cells(23, nColSubInd + 59)) = matSubInd

'**********************
'LOG file
'**********************

w_ns.Range(w_ns.cells(2, 2), w_ns.cells(1 + N, 1 + nfact)) = fact
'w_ns.Range(w_ns.Cells(2, 1 + nfact + 10), w_ns.Cells(1 + n, nfact + nfact + 11)) = isin_mat
'w_ns.Range(w_ns.Cells(2, nfact + 3), w_ns.Cells(61, nfact + 3 + nfact - 1)) = ISIN_sorted
w_ns.Range(w_ns.cells(2, nfact + 3), w_ns.cells(61, nfact + 3 + nfact - 1)) = ISIN_sorted
w_ns.Range(w_ns.cells(1, 2), w_ns.cells(1, nfact + 1)) = key
w_ns.Range(w_ns.cells(1, nfact + 3), w_ns.cells(1, nfact + 2 + nfact)) = key


'**********************
'End LOG file
'**********************



With w_c.Range(w_c.cells(3, nColSubInd), w_c.cells(3 + 20, nColSubInd + 59))
    .Columns.ColumnWidth = 15
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
End With
w_c.Range(w_c.cells(3, nColSubInd), w_c.cells(3, nColSubInd + 59)).Borders(xlEdgeBottom).LineStyle = xlContinuous
w_c.Range(w_c.cells(4, nColSubInd), w_c.cells(4, nColSubInd + 59)).Borders(xlEdgeBottom).LineStyle = xlDouble
'formatting sheets
'sum col
With w_c.cells(3, sumCol)
    .Font.Bold = True
    .Font.Color = rgb(255, 255, 255)
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Interior.ColorIndex = 32
    .HorizontalAlignment = xlCenter

End With

With w_c.Range(w_c.cells(4, sumCol), w_c.cells(3 + N, sumCol))
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
'    .Interior.Color = RGB(246, 194, 102)
End With


'ranked col
With w_c.Range(w_c.cells(3, rankCol), w_c.cells(3, rankCol + 4))
    .Font.Bold = True
    .Font.Color = rgb(255, 255, 255)
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Interior.ColorIndex = 45
    .HorizontalAlignment = xlCenter
    .Columns.AutoFit
End With

w_c.Range(w_c.cells(3, rankCol), w_c.cells(3 + N, rankCol)).HorizontalAlignment = xlCenter
w_c.Range(w_c.cells(3, rankCol + 1), w_c.cells(3, rankCol + 1)).Columns.ColumnWidth = 16
w_c.Range(w_c.cells(3, rankCol + 2), w_c.cells(3, rankCol + 2)).Columns.ColumnWidth = 35
w_c.Range(w_c.cells(3, rankCol + 3), w_c.cells(3, rankCol + 3)).Columns.ColumnWidth = 35
w_c.Range(w_c.cells(3, rankCol + 4), w_c.cells(3, rankCol + 4)).Columns.ColumnWidth = 16

With w_c.Range(w_c.cells(4, rankCol), w_c.cells(3 + N, rankCol + 4))
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Font.Bold = True
'    .Columns.ColumnWidth = 17
    .HorizontalAlignment = xlLeft
'    .Interior.Color = RGB(246, 194, 102)
End With
w_c.Range(w_c.cells(3, rankCol), w_c.cells(3 + N, rankCol)).HorizontalAlignment = xlCenter

'tabella conta col
With w_c.Range(w_c.cells(3, 9), w_c.cells(3, 8 + nfact))
    .Font.Bold = True
    .Font.Color = rgb(255, 255, 255)
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Interior.ColorIndex = 3
    '.HorizontalAlignment = xlCenter
End With

With w_c.Range(w_c.cells(4, 9), w_c.cells(3 + N, 8 + nfact))
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
'    .Interior.Color = RGB(246, 194, 102)
End With


'third row heigth
w_c.Rows(3).RowHeight = 35

With w_c.Rows(3)
    .VerticalAlignment = xlCenter
    .HorizontalAlignment = xlCenter
    .WrapText = True
    .ShrinkToFit = True
End With
    
SecondsElapsed = Round(Timer - StartTime, 2)

With w_c.Range(w_c.cells(3, 1), w_c.cells(3, 8))
    .Font.Bold = True
    .Font.Color = rgb(255, 255, 255)
    .Interior.ColorIndex = 1
End With

'highlight stocks in portfolio
'Set stockArray = Sheets("PTF EQUITY").Range(Sheets("PTF EQUITY").Cells(8, 4), Sheets("PTF EQUITY").Cells(8, 4).End(xlDown))
nEQT1 = Sheets("PTF EQUITY").Range(Sheets("PTF EQUITY").cells(8, 4), Sheets("PTF EQUITY").cells(8, 4).End(xlDown)).Count

For i = 1 To N
    For J = 1 To nEQT1
        If w_c.cells(3 + i, rankCol + 1) = Sheets("PTF EQUITY").cells(7 + J, 4) Then
           w_c.Range(w_c.cells(3 + i, rankCol), w_c.cells(3 + i, rankCol + 4)).Font.Color = rgb(255, 0, 0)
        End If
    Next J
Next i

w_c.Columns("I:Z").NumberFormat = "0"
Debug.Print SecondsElapsed
Application.Calculation = xlCalculationAutomatic
End Sub

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

Function calcGrowth(mat1 As Range, mat2 As Range) As Variant

Dim i As Integer
Dim m1() As Variant
Dim m2() As Variant
Dim res() As Variant

N = mat1.Rows.Count

Set mat1 = mat1
Set mat2 = mat2

ReDim m1(1 To N, 1)
ReDim m2(1 To N, 1)
ReDim res(1 To N, 1)

m1 = mat1.Value
m2 = mat2.Value

For i = LBound(m1, 1) To UBound(m1, 1)


    res(i, 1) = mat2(i, 1) / mat1(i, 1) - 1
    
Next i

calcGrowth = res
End Function

Sub calc()
Dim i As Integer
Dim m1() As Variant
Dim m2() As Variant
Dim res() As Variant

Set mat1 = Range("B5:B8")
Set mat2 = Range("C5:C8")


N = mat1.Rows.Count

ReDim m1(1 To N, 1)
ReDim m2(1 To N, 1)
ReDim res(1 To N, 1)

m1 = mat1.Value
m2 = mat2.Value


'Debug.Print UBound(m1, 1)

For i = LBound(m1, 1) To UBound(m1, 1)

    res(i, 1) = mat2(i, 1) / mat1(i, 1) - 1
   ' Debug.Print res(i, 1)

Next i
End Sub

Function give_title(sh As String, nrow As Integer, ncol As Integer) As String
    If Not (IsNumeric(Sheets(sh).cells(nrow, ncol)) Or Sheets(sh).cells(nrow, ncol) = "") Then
        give_title = Sheets(sh).cells(nrow, ncol).Text
    Else:
        nrow = nrow - 2
        If Sheets(sh).cells(nrow, ncol).Text = "" Then
           give_title = give_title(sh, nrow, ncol - 1)
        Else
            give_title = give_title(sh, nrow - 1, ncol)
        End If
    End If
End Function

Sub Test()
a = Sheets("Monitor Azioni").Range("ad5").MergeArea.Address
a = Split(a, ":")
Debug.Print a(0)
End Sub
Sub DeleteCells()
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Sheets("Tab").UsedRange.Delete
    Application.ScreenUpdating = True
End Sub

Sub DeleteCellsLOG()
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Sheets("log").UsedRange.Delete
    Application.ScreenUpdating = True
End Sub
