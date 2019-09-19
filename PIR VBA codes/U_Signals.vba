Sub generateMASignal()
Dim ws_d As Worksheet
Dim priceMat() As Variant

m_ma = 12
n_ma = 20
k_tbr = 10

Set ws_d = Sheets("Data")
nStock = ws_d.cells(7, 3)
nDate = ws_d.cells(7, 264)

a = 0
For i = 1 To nDate
    If ws_d.cells(7 + i, 266) = "" Then
        a = a + 1
    End If
Next i

StartDate = Format(CDate("07/03/2017"), "dd/mm/yyyy")
EndDate = CDate(Application.WorksheetFunction.WorkDay(Now(), -1))
diff_date = Application.WorksheetFunction.NetworkDays(StartDate, EndDate)

ReDim priceMat(1 To diff_date, 1 To nStock)

Set data_p = ws_d.Range(ws_d.cells(8, 266), ws_d.cells(7 + diff_date, 266 + nStock - 1))
priceMat = data_p.Value


ReDim ShortMa(1 To 2, 1 To nStock)
'Debug.Print UBound(priceMat, 1)
For i = 1 To nStock
    tshort_MA = 0
    t_1short_MA = 0
    For J = UBound(priceMat, 1) To (UBound(priceMat, 1) - m_ma + 1) Step -1
        tshort_MA = tshort_MA + priceMat(J, i)
        t_1short_MA = t_1short_MA + priceMat(J - 1, i)
    Next J
    tshort_MA = tshort_MA / m_ma
    t_1short_MA = t_1short_MA / m_ma
    
    ShortMa(1, i) = tshort_MA
    ShortMa(2, i) = t_1short_MA
'    Debug.Print ShortMa(1, i); ShortMa(2, i)
Next i

ReDim LongMa(1 To 2, 1 To nStock)

For i = 1 To nStock
    tlong_MA = 0
    t_1long_MA = 0
    For J = UBound(priceMat, 1) To (UBound(priceMat, 1) - n_ma + 1) Step -1
        tlong_MA = tlong_MA + priceMat(J, i)
        t_1long_MA = t_1long_MA + priceMat(J - 1, i)
    Next J
    tlong_MA = tlong_MA / n_ma
    t_1long_MA = t_1long_MA / n_ma
    
    LongMa(1, i) = tlong_MA
    LongMa(2, i) = t_1long_MA
    
Next i

ReDim signalMA(1 To 1, 1 To nStock)
For i = 1 To nStock
    MAO_t_sig = Sgn(ShortMa(1, i) - LongMa(1, i))
    MAO_t1_sig = Sgn(ShortMa(2, i) - LongMa(2, i))
    If MAO_t_sig - MAO_t1_sig = 2 Then
        signalMA(1, i) = "BUY"
    ElseIf MAO_t_sig - MAO_t1_sig = -2 Then
        signalMA(1, i) = "SELL"
    Else
        signalMA(1, i) = "Neutral"
    End If
Next i

'''Trading range breakout

'ReDim vectTBR(1 To 3, 1 To nStock)
ReDim signalTBR(1 To 1, 1 To nStock)
For i = 1 To nStock
    ReDim tempV(1 To k_tbr, 1 To 1)
    indx = 1
    For J = UBound(priceMat, 1) - 1 To (UBound(priceMat, 1) - k_tbr) Step -1
        tempV(indx, 1) = priceMat(J, i)
        indx = indx + 1
    Next J
    max_temp = Application.Max(tempV)
    min_temp = Application.Min(tempV)
'    vectTBR(1, i) = max_temp
'    vectTBR(2, i) = min_temp
'    vectTBR(3, i) = priceMat(UBound(priceMat, 1), i)
   tbr = (priceMat(UBound(priceMat, 1), i) > max_temp) - (priceMat(UBound(priceMat, 1), i) < min_temp)
    If tbr = 1 Then
       signalTBR(1, i) = "BUY"
    ElseIf tbr = -1 Then
        signalTBR(1, i) = "SELL"
    Else
        signalTBR(1, i) = "Neutral"
    End If
Next i

ws_d.Range(ws_d.cells(9, 406), ws_d.cells(9 + nStock - 1, 406)) = Application.Transpose(signalMA)
ws_d.Range(ws_d.cells(9, 407), ws_d.cells(9 + nStock - 1, 407)) = Application.Transpose(signalTBR)
    
End Sub