'Option Base 1
'Sub Screening_sorting()
'
'Dim arr As Variant
'Dim vect_val As Variant
'Dim vect_isin As Variant
'Dim ISIN_sorted(1 To 20, 1) As Variant
'Dim All_isin_sorted(1 To 20, 1 To 58) As Variant
'Dim All_ones_or_zeros() As Variant
'For niter = 1 To 58
'    If Sheets("Monitor Azioni").Cells(7, 9 + niter) = "" Or Sheets("Monitor Azioni").Cells(6, 9 + niter) = "TICKER" Or Sheets("Monitor Azioni").Cells(6, 9 + niter) = "DESCRIZIONE" Then GoTo NextIteration
'    'Debug.Print Cells(3, 11 + niter)
'    'Debug.Print niter + 9
''    Sheets("Array").Cells(4, 1 + niter) = Cells(3, 9 + niter)
'    Set RNG1 = Sheets("Monitor Azioni").Range(Sheets("Monitor Azioni").Cells(7, 9 + niter), Sheets("Monitor Azioni").Cells(7, 9 + niter).End(xlDown))
'    vect_val = RNG1.Value
'
'    Set isin = Sheets("Monitor Azioni").Range(Sheets("Monitor Azioni").Cells(7, 2), Sheets("Monitor Azioni").Cells(7, 2).End(xlDown))
'    vect_isin = isin.Value
'    ReDim arr(1 To RNG1.Rows.Count, 1 To 2)
'
'
'
'    For i = LBound(arr, 1) To UBound(arr, 1)
'        If IsNumeric(vect_val(i, 1)) Then
'        arr(i, 1) = vect_val(i, 1)
'        Else
'        arr(i, 1) = -9999
'        End If
'        arr(i, 2) = vect_isin(i, 1)
'    Next i
'
'
'    QuicksortD arr, LBound(arr), UBound(arr), 1
'
'    'For i = LBound(arr, 1) To UBound(arr, 1)
'    '   Debug.Print arr(i, 1);
'    '   Debug.Print arr(i, 2)
'    'Next i
'
'    For i = 1 To 20
'        All_isin_sorted(i, niter) = arr(i, 2)
'    Next i
'    'Debug.Print Isin_sorted(1, 1)
'NextIteration:
'Next niter
'
''Sheets("Array").Range("B5:BF23").Value = All_isin_sorted
'Sheets("Monitor Azioni").Range(Sheets("Monitor Azioni").Cells(4, 10), Sheets("Monitor Azioni").Cells(6, 66)).Copy
'Sheets("Tab").Range("J1").PasteSpecial
'Sheets("Tab").Range("J1").PasteSpecial xlPasteValuesAndNumberFormats
'Application.CutCopyMode = False
'
'ReDim All_ones_or_zeros(1 To RNG1.Rows.Count, 1 To 58)
'For j = 1 To 58
'        On Error Resume Next
'    Err.Clear
'    For i = 1 To 20
'
'        a = Application.WorksheetFunction.Match(All_isin_sorted(i, j), Sheets("Tab").Range("B4:B1000"), 0)
'        If Err.Number = 0 Then
''           Sheets("Tab").Cells(a, 9 + j) = 1
''            Sheets("Tab").Cells(a, 9 + j).NumberFormat = "0"
'            All_ones_or_zeros(a, j) = 1
'        Else: GoTo nit
'        End If
'Next i
'nit:
'Next j
'
'Sheets("Tab").Range("J4:BN263").Value = All_ones_or_zeros
'
'
''col_index_array = Array(12, 16, 13, 19, 10, 25, 30, 35, 41, 45, 51, 56, 22, 59, 66, 63)
''Start = 10
''    While Start < 66
'''        For i = LBound(col_index_array) To UBound(col_index_array)
'''             Debug.Print col_index_array(i)
'''         Next i
''            If IsInArray(Start, col_index_array) Then
''            Else: Sheets("Tab").Columns(Start).Delete shift:=xlShiftToLeft
''            End If
''            Start = Start + 1
''            Wend
'End Sub
'Sub QuicksortD(ary, LB, UB, ref)
'Dim M As Variant, temp
'Dim i As Long, ii As Long, iii As Integer
'i = UB
'ii = LB
'M = ary(Int((LB + UB) / 2), ref)
'Do While ii <= i
'    Do While ary(ii, ref) > M
'        ii = ii + 1
'    Loop
'    Do While ary(i, ref) < M
'        i = i - 1
'    Loop
'    If ii <= i Then
'        For iii = LBound(ary, 2) To UBound(ary, 2)
'            temp = ary(ii, iii): ary(ii, iii) = ary(i, iii)
'            ary(i, iii) = temp
'        Next
'        ii = ii + 1: i = i - 1
'    End If
'Loop
'If LB < i Then QuicksortD ary, LB, i, ref
'If ii < UB Then QuicksortD ary, ii, UB, ref
'End Sub
'Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
''DEVELOPER: Ryan Wells (wellsr.com)
''DESCRIPTION: Function to check if a value is in an array of values
''INPUT: Pass the function a value to search for and an array of values of any data type.
''OUTPUT: True if is in array, false otherwise
'Dim element As Variant
'On Error GoTo IsInArrayError: 'array is empty
'    For Each element In arr
'        If element = valToBeFound Then
'            IsInArray = True
'            Exit Function
'        End If
'    Next element
'Exit Function
'IsInArrayError:
'On Error GoTo 0
'IsInArray = False
'End Function
''Sub doItManually()
''
''Set isin = Range(Cells(7, 2), Cells(7, 2).End(xlDown))
''
''Set YTD = Sheets("Monitor Azioni").Range(Cells(7, 12), Cells(7, 12).End(xlDown))
''Set MTD = Sheets("Monitor Azioni").Range(Cells(7, 13), Cells(7, 13).End(xlDown))
''Set D5 = Sheets("Monitor Azioni").Range(Cells(7, 14), Cells(7, 14).End(xlDown))
''Set DYL = Sheets("Monitor Azioni").Range(Cells(7, 15), Cells(7, 15).End(xlDown))
''Set M3 = Sheets("Monitor Azioni").Range(Cells(7, 16), Cells(7, 16).End(xlDown))
''Set M6 = Sheets("Monitor Azioni").Range(Cells(7, 17), Cells(7, 17).End(xlDown))
''Set YR = Sheets("Monitor Azioni").Range(Cells(7, 18), Cells(7, 18).End(xlDown))
''
''Set vol = Sheets("Monitor Azioni").Range(Cells(7, 19), Cells(7, 19).End(xlDown))
''Set cap = Sheets("Monitor Azioni").Range(Cells(7, 10), Cells(7, 10).End(xlDown))
''Set pot = Sheets("Monitor Azioni").Range(Cells(7, 25), Cells(7, 25).End(xlDown))
''
''Set rev_next = Sheets("Monitor Azioni").Range(Cells(7, 30), Cells(7, 30).End(xlDown))
''Set rev_last = Sheets("Monitor Azioni").Range(Cells(7, 29), Cells(7, 29).End(xlDown))
''Set ebd = Sheets("Monitor Azioni").Range(Cells(7, 35), Cells(7, 35).End(xlDown))
''Set eps = Sheets("Monitor Azioni").Range(Cells(7, 41), Cells(7, 41).End(xlDown))
''Set div = Sheets("Monitor Azioni").Range(Cells(7, 45), Cells(7, 45).End(xlDown))
''
''Set PE = Sheets("Monitor Azioni").Range(Cells(7, 51), Cells(7, 51).End(xlDown))
'''set PB = Sheets("Monitor Azioni").Range(Cells(7, 54), Cells(7, 54).End(xlDown))
''Set EV_EBTDA = Sheets("Monitor Azioni").Range(Cells(7, 56), Cells(7, 56).End(xlDown))
''Set float = Sheets("Monitor Azioni").Range(Cells(7, 22), Cells(7, 22).End(xlDown))
''
''Set ROE = Sheets("Monitor Azioni").Range(Cells(7, 59), Cells(7, 59).End(xlDown))
''Set FCF = Sheets("Monitor Azioni").Range(Cells(7, 66), Cells(7, 66).End(xlDown))
''Set G = Sheets("Monitor Azioni").Range(Cells(7, 63), Cells(7, 63).End(xlDown))
''
''End Sub
'
'Sub forloopo()
'
'col_index_array = Array(12, 16, 13, 19, 10, 25, 30, 35, 41, 45, 51, 56, 22, 59, 66, 63)
'Start = 10
'    While Start < 66
''        For i = LBound(col_index_array) To UBound(col_index_array)
''             Debug.Print col_index_array(i)
''         Next i
'            If IsInArray(Start, col_index_array) Then
'            Else: Sheets("Tab").Columns(Start).Delete shift:=xlShiftToLeft
'            End If
'            Start = Start + 1
'            Wend
'
'
'End Sub
'Sub UnMergeFill()
'
'Dim cell As Range, joinedCells As Range
'
'For Each cell In ThisWorkbook.ActiveSheet.UsedRange
'    If cell.MergeCells Then
'        Set joinedCells = cell.MergeArea
'        cell.MergeCells = False
'        joinedCells.Value = cell.Value
'    End If
'Next
'
'End Sub
'Sub newdelcol()
'    With Sheets("Tab")
'        For currentColumn = .UsedRange.Columns.Count To 1 Step -1
'
'            columnHeading = .UsedRange.Columns.Column
'Debug.Print currentColumn
'            'CHECK WHETHER TO KEEP THE COLUMN
'            Select Case currentColumn
'
'                Case 66, 63, 59, 56, 51, 45, 35, 41, 30, 25, 22, 19, 16, 13, 12, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1
'                    'Do nothing
'                Case Else
'
'                     Sheets("Tab").Columns(columnHeading).Delete shift:=xlShiftToLeft
'            End Select
'        Next
'    End With
'
'
'End Sub
'Sub deb()
'With Worksheets("Tab")
'a = .Columns(5).Column
'End With
'Debug.Print a
'End Sub
'