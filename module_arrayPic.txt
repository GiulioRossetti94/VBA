'======================================================================================================================
Sub importArray()

    n_pic = Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.ListCount
    
    If Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.Value = "RANDOM" Then
        Index_pic = WorksheetFunction.RandBetween(1, n_pic - 1)
        file_txt_name = "array_TXT" & Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.List(Index_pic) & ".txt"
    Else
        file_txt_name = "array_TXT" & Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.Value & ".txt"
    End If
    
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Scripts Python\PIC_ARRAY\" & file_txt_name _
        , Destination:=Range("$BE$28"))
        '.Name = "array_TXT_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 850
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = True
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
        2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Columns("BE:CR").Select
    Selection.ColumnWidth = 2.43
    Range("L41").Activate
End Sub

'======================================================================================================================

Sub HEXtoCOL()
'''OLD CODE go down for new one
For I = 1 To 40
    For J = 1 To 40
        color_txt = cells(27 + I, 56 + J)
        
        'Debug.Print color_txt

            red = Int(WorksheetFunction.Hex2Dec(Left(color_txt, 2)))
            green = Int(WorksheetFunction.Hex2Dec(Mid(color_txt, 4, 2)))
            blue = Int(WorksheetFunction.Hex2Dec(Right(color_txt, 2)))


        
        If ActiveSheet.CheckBox5.Value = True Then
            green = 255 - green
            red = 255 - red
            blue = 255 - blue
        End If
        
        If ActiveSheet.CheckBox6.Value = True Then
            If green = 0 And red = 0 And blue = 0 Then
                green = 255
                red = 255
                blue = 255
            ElseIf green = 255 And red = 255 And blue = 255 Then
                green = 0
                red = 0
                blue = 0
            End If
        End If
        
        If ActiveSheet.CheckBox7.Value = True Then
            avg = Int((Int(green) + Int(red) + Int(blue)) / 3)
            If ActiveSheet.CheckBox5.Value = True Then avg = 255 - avg
         
            cells(27 + I, 12 + J).Interior.Color = rgb(avg, avg, avg)
            
        Else
            cells(27 + I, 12 + J).Interior.Color = rgb(red, green, blue)

        End If

       
    Next J
Next I
End Sub

'======================================================================================================================

Sub PythonScript()
path = Chr(34) & "Y:\\Mobiliare\\08" & " " & "Finint" & " " & "Economia" & " " & "Reale" & " " & "Italia\\01_Front" & " " & "Office\\02" & " " & "Gestione\\Scripts" & " " & "Python\\PIC_ARRAY\\pixel_ex.py" & Chr(34)
Debug.Print path

run_anaconda_prompt = "C:/ProgramData/Anaconda3/Scripts/activate.bat C:\ProgramData\Anaconda3"
RetVal = Shell("C:/ProgramData/Anaconda3/Scripts/activate.bat C:\ProgramData\Anaconda3 python")


End Sub

'======================================================================================================================

Sub newHEXtoCOL()
'
Dim color_index(1 To 3) As Variant
Dim color_array(1 To 3) As Variant
Dim color_name(1 To 3) As Variant

color_name(1) = "Red"
color_name(2) = "Green"
color_name(3) = "Blue"
'
color_index(1) = 1
color_index(2) = 2
color_index(3) = 3

ShuffleArrayInPlace color_index

Worksheets("SCARICHI BLOOMBERG").OLEObjects("TextBox1").Object.Value = ""
'Debug.Print color_name(color_index(1)), color_name(color_index(2)), color_name(color_index(3))

For I = 1 To 40
    For J = 1 To 40
        color_txt = cells(27 + I, 56 + J)
        
        num_array = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f")
        char1 = LCase(Mid(color_txt, 1, 1))
        char2 = LCase(Mid(color_txt, 2, 1))
        char3 = LCase(Mid(color_txt, 3, 1))
        char4 = LCase(Mid(color_txt, 4, 1))
        char5 = LCase(Mid(color_txt, 5, 1))
        char6 = LCase(Mid(color_txt, 6, 1))
        
        For k = 0 To 15
            If (char1 = num_array(k)) Then position1 = k
            If (char2 = num_array(k)) Then position2 = k
            If (char3 = num_array(k)) Then position3 = k
            If (char4 = num_array(k)) Then position4 = k
            If (char5 = num_array(k)) Then position5 = k
            If (char6 = num_array(k)) Then position6 = k
        Next k
        
        red = position1 * 16 + position2
        green = position3 * 16 + position4
        blue = position5 * 16 + position6
        
       
        If ActiveSheet.CheckBox5.Value = True Then
            green = 255 - green
            red = 255 - red
            blue = 255 - blue
        End If
        
        If ActiveSheet.CheckBox6.Value = True Then
            If green < 50 And red < 50 And blue < 50 Then
                green = 255
                red = 255
                blue = 255
            ElseIf green > 200 And red > 200 And blue > 200 Then
                green = 0
                red = 0
                blue = 0
            End If
        End If
        
        If ActiveSheet.CheckBox7.Value = True Then
            avg = Int((Int(green) + Int(red) + Int(blue)) / 3)
            If ActiveSheet.CheckBox5.Value = True Then avg = 255 - avg
         
            cells(27 + I, 12 + J).Interior.Color = rgb(avg, avg, avg)
            
        ElseIf ActiveSheet.CheckBox10.Value = True Then
                'cells(27 + i, 12 + j).Interior.Color = RGB(blue, red, green)
                Dim color_array_random(1 To 3) As Variant
                
                color_array_random(1) = red
                color_array_random(2) = green
                color_array_random(3) = blue
                
                ShuffleArrayInPlace color_array_random
                cells(27 + I, 12 + J).Interior.Color = rgb(color_array_random(1), color_array_random(2), color_array_random(3))

         ElseIf ActiveSheet.CheckBox9.Value = True Then
                color_array(1) = red
                color_array(2) = green
                color_array(3) = blue

                cells(27 + I, 12 + J).Interior.Color = rgb(color_array(color_index(1)), color_array(color_index(2)), color_array(color_index(3)))
        Else
        
                cells(27 + I, 12 + J).Interior.Color = rgb(red, green, blue)

        
        End If
       
    Next J
Next I

If ActiveSheet.CheckBox9.Value = True Then
Worksheets("SCARICHI BLOOMBERG").OLEObjects("TextBox1").Object.Value = "RGB: " & color_name(color_index(1)) & " - " & color_name(color_index(2)) & " - " & color_name(color_index(3))
End If

End Sub

'======================================================================================================================

Sub NewFastImport()
'not working'
    n_pic = Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.ListCount
    
    If Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.Value = "RANDOM" Then
        Index_pic = WorksheetFunction.RandBetween(1, n_pic - 1)
        file_txt_name = "array_TXT" & Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.List(Index_pic) & ".txt"
    Else
        file_txt_name = "array_TXT" & Worksheets("SCARICHI BLOOMBERG").OLEObjects("ComboBox1").Object.Value & ".txt"
    End If
    
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;Y:\Mobiliare\08 Finint Economia Reale Italia\01_Front Office\02 Gestione\Scripts Python\PIC_ARRAY\" & file_txt_name _
        , Destination:=Range("$BE$28"))
        .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
                                    2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
    End With
    Columns("BE:CR").Select
    Selection.ColumnWidth = 2.43
    Range("L41").Activate

End Sub

'======================================================================================================================

Sub ShuffleArrayInPlace(InArray() As Variant)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShuffleArrayInPlace
' This shuffles InArray to random order, randomized in place.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim L As Long
    Dim Temp As Variant
    Dim J As Long
    
    Randomize
    L = UBound(InArray) - LBound(InArray) + 1
    For N = LBound(InArray) To UBound(InArray)
        J = Int((UBound(InArray) - LBound(InArray) + 1) * Rnd + LBound(InArray))
        If N <> J Then
            Temp = InArray(N)
            InArray(N) = InArray(J)
            InArray(J) = Temp
        End If
    Next N
End Sub

'======================================================================================================================

Sub delNamedRange()

For Each NR In ActiveWorkbook.Names
    If InStr(NR.Value, "='SCARICHI BLOOMBERG'!$BE$28:$CR$67") Then NR.Delete
Next
End Sub
'======================================================================================================================
Sub ColorOrdering()
Dim colMat(1 To 40, 1 To 40) As Variant
Dim orderedMat(1 To 40, 1 To 40) As Variant
Dim orderedMatFinal(1 To 40, 1 To 40) As Variant
Dim slice As Variant

For I = 1 To 40
    For J = 1 To 40
        colMat(I, J) = cells(27 + I, 12 + J).DisplayFormat.Interior.Color
    Next J
Next I

For I = 1 To 40
    slice = Application.WorksheetFunction.Index(colMat, 0, I)
    QuicksortD slice, LBound(slice), UBound(slice), 1
    For J = 1 To 40
    'Range(cells(28, 100 + i), cells(67, 100 + i)) = slice
    'Debug.Print slice(i, 1)
        orderedMat(J, I) = slice(J, 1)
    Next J
Next I

For I = 1 To 40
    slice = Application.Transpose(Application.WorksheetFunction.Index(orderedMat, I, 0))
    
    QuicksortD slice, LBound(slice), UBound(slice), 1
    For J = 1 To 40
    'Range(cells(28, 100 + i), cells(67, 100 + i)) = slice

        orderedMatFinal(J, I) = slice(J, 1)
    Next J
Next I

For I = 1 To 40
    For J = 1 To 40
        
        clr = orderedMatFinal(I, J)
        red = clr And 255
        green = clr \ 256 And 255
        blue = clr \ 256 ^ 2 And 255

        cells(27 + I, 12 + J).Interior.Color = rgb(red, green, blue)

    Next J
Next I

End Sub

'======================================================================================================================

Sub fake()
Dim colMat(1 To 40, 1 To 40) As Variant
For I = 1 To 40
    For J = 1 To 40
        colMat(I, J) = cells(27 + I, 12 + J).DisplayFormat.Interior.Color
        Debug.Print colMat(I, J)
    Next J
Next I
End Sub
'======================================================================================================================

Sub QuicksortD(ary, LB, UB, ref)
Dim M As Variant, Temp
Dim I As Long, ii As Long, iii As Integer
I = UB
ii = LB
M = ary(Int((LB + UB) / 2), ref)
Do While ii <= I
    Do While ary(ii, ref) > M
        ii = ii + 1
    Loop
    Do While ary(I, ref) < M
        I = I - 1
    Loop
    If ii <= I Then
        For iii = LBound(ary, 2) To UBound(ary, 2)
            Temp = ary(ii, iii): ary(ii, iii) = ary(I, iii)
            ary(I, iii) = Temp
        Next
        ii = ii + 1: I = I - 1
    End If
Loop
If LB < I Then QuicksortD ary, LB, I, ref
If ii < UB Then QuicksortD ary, ii, UB, ref
End Sub

'======================================================================================================================

Function HexToRGB(color_txt As String) As Variant

num_array = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f")
char1 = LCase(Mid(color_txt, 1, 1))
char2 = LCase(Mid(color_txt, 2, 1))
char3 = LCase(Mid(color_txt, 3, 1))
char4 = LCase(Mid(color_txt, 4, 1))
char5 = LCase(Mid(color_txt, 5, 1))
char6 = LCase(Mid(color_txt, 6, 1))

For k = 0 To 15
    If (char1 = num_array(k)) Then position1 = k
    If (char2 = num_array(k)) Then position2 = k
    If (char3 = num_array(k)) Then position3 = k
    If (char4 = num_array(k)) Then position4 = k
    If (char5 = num_array(k)) Then position5 = k
    If (char6 = num_array(k)) Then position6 = k
Next k

red = position1 * 16 + position2
green = position3 * 16 + position4
blue = position5 * 16 + position6


HexToRGB = Array(red, green, blue)
End Function

'======================================================================================================================

Function HexCode(Cell As Range) As String
    HexCode = Right("000000" & Hex(Cell.Interior.Color), 6)
    HexCode = Right(HexCode, 2) & Mid(HexCode, 3, 2) & Left(HexCode, 2)
End Function
'======================================================================================================================
Function RGBtoHSV(rgb As Variant) As Variant
Dim hsv(1 To 3) As Variant
red = rgb(1)
green = rgb(2)
blue = rgb(3)


r = red / 255
g = green / 255
b = blue / 255
max_val = Application.WorksheetFunction.Max(r, g, b)
min_val = Application.WorksheetFunction.Min(r, g, b)
c = max_val - min_val

If c = 0 Then
    hue = 0
Else
    If r = max_val Then
        segment = (g - b) / c
        shift = 0 / 60
        If segment < 0 Then shift = 360 / 60
        hue = segment + shift
    ElseIf g = max_val Then
        segment = (b - r) / c
        shift = 120 / 60
        hue = segment + shift
    ElseIf b = max_val Then
        segment = (r - g) / c
        shift = 240 / 60
        hue = segment + shift
    End If
End If

hue = hue * 60
If max_val = 0 Then
    sat = 0
Else
    sat = c / max_val
End If
val_l = max_val
        
hsv(1) = hue
hsv(2) = sat
hsv(3) = val_l

RGBtoHSV = hsv
End Function

'======================================================================================================================

'Function RGBtoHSL(rgb As Variant) As Variant
'Dim hsl(1 To 3) As Variant
'red = rgb(1)
'green = rgb(2)
'blue = rgb(3)
'
'
'r = red / 255
'g = green / 255
'b = blue / 255
'
'max_val = Application.WorksheetFunction.Max(r, g, b)
'min_val = Application.WorksheetFunction.Min(r, g, b)
'C = max_val - min_val
'
'If C = 0 Then
'    hue = 0
'Else
'    If r = max_val Then
'        segment = (g - b) / C
'        shift = 0 / 60
'        If segment < 0 Then shift = 360 / 60
'        hue = segment + shift
'    ElseIf g = max_val Then
'        segment = (b - r) / C
'        shift = 120 / 60
'        hue = segment + shift
'    ElseIf b = max_val Then
'        segment = (r - g) / C
'        shift = 240 / 60
'        hue = segment + shift
'    End If
'End If
'
'hue = hue * 60
'
'light = (max_val + min_val) / 2
'
'If max_val = 0 Then
'    sat = 0
'ElseIf min_val = 1 Then
'    sat = 0
'Else
'    sat = (max_val - light) / (Application.WorksheetFunction.Min(light, 1 - light))
'End If
'
'
'hsl(1) = hue
'hsl(2) = sat
'hsl(3) = light
'
'RGBtoHSL = hsl
'End Function

'======================================================================================================================

Function RGBtoHSL(rgb As Variant) As Variant
Dim hsl(1 To 3) As Variant
red = rgb(1)
green = rgb(2)
blue = rgb(3)


r = red / 255
g = green / 255
b = blue / 255

max_val = Application.WorksheetFunction.Max(r, g, b)
min_val = Application.WorksheetFunction.Min(r, g, b)
c = max_val - min_val

If c = 0 Then
    hue = 0
Else
    If r = max_val Then
        hue = 60 * (0 + (g - b) / c)
    ElseIf g = max_val Then
        hue = 60 * (2 + (b - r) / c)
    ElseIf b = max_val Then
        hue = 60 * (4 + (r - g) / c)
    End If
End If

If hue < 0 Then hue = hue + 360

light = (max_val + min_val) / 2

If max_val = 0 Then
    sat = 0
ElseIf min_val = 1 Then
    sat = 0
Else
    sat = (max_val - light) / (Application.WorksheetFunction.Min(light, 1 - light))
End If

        
hsl(1) = hue
hsl(2) = sat
hsl(3) = light

RGBtoHSL = hsl
End Function

'======================================================================================================================

Function HSLtoRGB(hsl As Variant) As Variant
Dim rgb(1 To 3) As Variant
h = hsl(1)
s = hsl(2)
L = hsl(3)

c = (1 - Abs(2 * L - 1)) * s
h = h / 60
x = c - (1 - Abs(h Mod 2 - 1))

If h >= 0 And h <= 1 Then
    r = c
    g = x
    b = 0
ElseIf h >= 1 And h <= 2 Then
    r = x
    g = c
    b = 0
ElseIf h >= 2 And h <= 3 Then
    r = 0
    g = c
    b = x
ElseIf h >= 3 And h <= 4 Then
    r = 0
    g = x
    b = c
ElseIf h >= 4 And h <= 5 Then
    r = x
    g = 0
    b = c
ElseIf h >= 5 And h <= 6 Then
    r = c
    g = 0
    b = x
Else
    r = 0
    g = 0
    b = 0
End If
M = L - c / 2

rgb(1) = (r + M) * 255
rgb(2) = (g + M) * 255
rgb(3) = (b + M) * 255
HSLtoRGB = rgb
End Function

'======================================================================================================================

Function RGBtoLight(rgb As Variant) As Integer
red = rgb(1)
green = rgb(2)
blue = rgb(3)

RGBtoLight = 0.2126 * red + 0.7152 * green + 0.0722 * blue
End Function

'======================================================================================================================

Sub BubbleSort(arr)
  Dim strTemp As String
  Dim I As Long
  Dim J As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(arr)
  lngMax = UBound(arr)
  For I = lngMin To lngMax - 1
    For J = I + 1 To lngMax
      If GetNumber(arr(I)) > GetNumber(arr(J)) Then
        strTemp = arr(I)
        arr(I) = arr(J)
        arr(J) = strTemp
      End If
    Next J
  Next I
End Sub

'======================================================================================================================


Sub ColorOrderingHSL()
'Sheets.Add.Name = "Temp"
'Set ws_t = Sheets("Temp")
Set ws_pic = Sheets("SCARICHI BLOOMBERG")

Application.DisplayAlerts = False
Dim colMat(1 To 40, 1 To 40) As Variant
Dim colors(1 To 40, 1 To 40) As Variant
Dim orderedMat(1 To 40, 1 To 40) As Variant
Dim orderedMatCol(1 To 40, 1 To 40) As Variant
Dim orderedMatFinal(1 To 40, 1 To 40) As Variant
Dim slice() As Variant
Dim slice2() As Variant
Dim v(1 To 3) As Variant
Dim arr(1 To 2) As Variant
Dim hsl(1 To 3) As Variant

For I = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + I, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        
        v(1) = red
        v(2) = green
        v(3) = blue
        
        colHSL = RGBtoHSL(v)
        h1 = Format((colHSL(1) / 360), "#0.000")
        s2 = Format((colHSL(2)), "#0.000")
        L3 = Format((colHSL(3)), "#0.000")
        
        
        colSortN = Round(5 * colHSL(3) + colHSL(2) * 2 + colHSL(1), 4)
        
        string_col = "H" & h1 & "S" & s2 & "L" & L3
        colors(I, J) = string_col
        
        colMat(I, J) = colSortN

    Next J
Next I


For I = 1 To 40
    slice = Application.Transpose(Application.WorksheetFunction.Index(colMat, 0, I))
    slice2 = Application.Transpose(Application.WorksheetFunction.Index(colors, 0, I))
    Call sort2(slice(), slice2())
    For J = 1 To 40

        orderedMat(J, I) = slice(J)
        orderedMatCol(J, I) = slice2(J)
    Next J
Next I

For I = 1 To 40
    slice = Application.WorksheetFunction.Index(orderedMat, I, 0)
    slice2 = Application.WorksheetFunction.Index(orderedMatCol, I, 0)
    Call sort2(slice(), slice2())
    For J = 1 To 40

        orderedMatFinal(J, I) = slice2(J)
    Next J
Next I

For I = 1 To 40
    For J = 1 To 40

        clr = orderedMatFinal(I, J)
        h = CDbl(Right(Left(clr, 6), 5))
        s = CDbl(Left(Right(clr, 11), 5))
        L = CDbl(Right(clr, 5))
        
        hsl(1) = h
        hsl(2) = s
        hsl(3) = L
        rgb_cols = HSLtoRGB(hsl)
        cells(27 + I, 12 + J).Interior.Color = rgb(rgb_cols(1), rgb_cols(2), rgb_cols(3))

    Next J
Next I
'ws_pic.Range("M28:AZ67") = orderedMatFinal
'ws_t.Range(ws_t.cells(1, 1), ws_t.cells(39, 39)) = orderedMatFinal
'Sheets("Temp").Delete
Application.DisplayAlerts = True
End Sub

'======================================================================================================================

Function getSortCol(s As Double) As Double

L = Len(s)
s = Left(s, Find("H", s) - 1)

getSortCol = s
End Function

'======================================================================================================================


Sub sort2(key() As Variant, other() As Variant)
Dim I As Long, J As Long, Low As Long
Dim Hi As Long, Temp As Variant
    Low = LBound(key)
    Hi = UBound(key)

    J = (Hi - Low + 1) \ 2
    Do While J > 0
        For I = Low To Hi - J
          If key(I) > key(I + J) Then
            Temp = key(I)
            key(I) = key(I + J)
            key(I + J) = Temp
            Temp = other(I)
            other(I) = other(I + J)
            other(I + J) = Temp
          End If
        Next I
        For I = Hi - J To Low Step -1
          If key(I) > key(I + J) Then
            Temp = key(I)
            key(I) = key(I + J)
            key(I + J) = Temp
            Temp = other(I)
            other(I) = other(I + J)
            other(I + J) = Temp
          End If
        Next I
        J = J \ 2
    Loop
End Sub

'======================================================================================================================
