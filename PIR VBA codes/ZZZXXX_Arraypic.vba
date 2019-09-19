Sub importArray()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'import txt file
'+++++++++++++++++++++++++++++++++++++++++++++++++++


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

Sub HEXtoCOL()
'''OLD CODE go down for new one
For i = 1 To 40
    For J = 1 To 40
        color_txt = cells(27 + i, 56 + J)
        
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
         
            cells(27 + i, 12 + J).Interior.Color = rgb(avg, avg, avg)
            
        Else
            cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)

        End If

       
    Next J
Next i




End Sub
Sub PythonScript()
path = Chr(34) & "Y:\\Mobiliare\\08" & " " & "Finint" & " " & "Economia" & " " & "Reale" & " " & "Italia\\01_Front" & " " & "Office\\02" & " " & "Gestione\\Scripts" & " " & "Python\\PIC_ARRAY\\pixel_ex.py" & Chr(34)
Debug.Print path

run_anaconda_prompt = "C:/ProgramData/Anaconda3/Scripts/activate.bat C:\ProgramData\Anaconda3"
RetVal = Shell("C:/ProgramData/Anaconda3/Scripts/activate.bat C:\ProgramData\Anaconda3 python")


End Sub


Sub newHEXtoCOL()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'convert colours from HEX to RGB
'+++++++++++++++++++++++++++++++++++++++++++++++++++

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

For i = 1 To 40
    For J = 1 To 40
        color_txt = cells(27 + i, 56 + J)
        
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
         
            cells(27 + i, 12 + J).Interior.Color = rgb(avg, avg, avg)
            
        ElseIf ActiveSheet.CheckBox10.Value = True Then
                'cells(27 + i, 12 + j).Interior.Color = RGB(blue, red, green)
                Dim color_array_random(1 To 3) As Variant
                
                color_array_random(1) = red
                color_array_random(2) = green
                color_array_random(3) = blue
                
                ShuffleArrayInPlace color_array_random
                cells(27 + i, 12 + J).Interior.Color = rgb(color_array_random(1), color_array_random(2), color_array_random(3))

         ElseIf ActiveSheet.CheckBox9.Value = True Then
                color_array(1) = red
                color_array(2) = green
                color_array(3) = blue

                cells(27 + i, 12 + J).Interior.Color = rgb(color_array(color_index(1)), color_array(color_index(2)), color_array(color_index(3)))
        Else
        
                cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)

        
        End If
       
    Next J
Next i

If ActiveSheet.CheckBox9.Value = True Then
Worksheets("SCARICHI BLOOMBERG").OLEObjects("TextBox1").Object.Value = "RGB: " & color_name(color_index(1)) & " - " & color_name(color_index(2)) & " - " & color_name(color_index(3))
End If


End Sub

Sub NewFastImport()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'Not working
'+++++++++++++++++++++++++++++++++++++++++++++++++++
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

Sub delNamedRange()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'every time a txt file is imported a new named range is create
'this sub deletes the named range associated with the importArray() sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++

For Each NR In ActiveWorkbook.Names
    If InStr(NR.Value, "='SCARICHI BLOOMBERG'!$BE$28:$CR$67") Then NR.Delete
Next
End Sub

Sub ColorOrdering()

'+++++++++++++++++++++++++++++++++++++++++++++++++++
'Sort color sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++


Dim colMat(1 To 40, 1 To 40) As Variant
Dim orderedMat(1 To 40, 1 To 40) As Variant
Dim orderedMatFinal(1 To 40, 1 To 40) As Variant
Dim slice As Variant

For i = 1 To 40
    For J = 1 To 40
        colMat(i, J) = cells(27 + i, 12 + J).DisplayFormat.Interior.Color
    Next J
Next i

For i = 1 To 40
    slice = Application.WorksheetFunction.Index(colMat, 0, i)
    QuicksortD slice, LBound(slice), UBound(slice), 1
    For J = 1 To 40
    'Range(cells(28, 100 + i), cells(67, 100 + i)) = slice
    'Debug.Print slice(i, 1)
        orderedMat(J, i) = slice(J, 1)
    Next J
Next i

For i = 1 To 40
    slice = Application.Transpose(Application.WorksheetFunction.Index(orderedMat, i, 0))
    
    QuicksortD slice, LBound(slice), UBound(slice), 1
    For J = 1 To 40
    'Range(cells(28, 100 + i), cells(67, 100 + i)) = slice

        orderedMatFinal(J, i) = slice(J, 1)
    Next J
Next i

For i = 1 To 40
    For J = 1 To 40
        
        clr = orderedMatFinal(i, J)
        red = clr And 255
        green = clr \ 256 And 255
        blue = clr \ 256 ^ 2 And 255

        cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)

    Next J
Next i

End Sub

Sub fake()
Dim colMat(1 To 40, 1 To 40) As Variant
For i = 1 To 40
    For J = 1 To 40
        colMat(i, J) = cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        Debug.Print colMat(i, J)
    Next J
Next i
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

Function HexToRGB(color_txt As String) As Variant

'+++++++++++++++++++++++++++++++++++++++++++++++++++
'convert colours from HEX to RGB. Put resulting rgb col,ors in an array
'+++++++++++++++++++++++++++++++++++++++++++++++++++

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
Function HexCode(cell As Range) As String
    HexCode = Right("000000" & Hex(cell.Interior.Color), 6)
    HexCode = Right(HexCode, 2) & Mid(HexCode, 3, 2) & Left(HexCode, 2)
End Function

Function RGBtoHSV(rgb As Variant) As Variant
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'convert colours from RGB to HSV. Not sure whether convertion works properly
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Dim hsv(1 To 3) As Variant
red = rgb(1)
green = rgb(2)
blue = rgb(3)


r = red / 255
g = green / 255
b = blue / 255
max_val = Application.WorksheetFunction.Max(r, g, b)
min_val = Application.WorksheetFunction.Min(r, g, b)
C = max_val - min_val

If C = 0 Then
    hue = 0
Else
    If r = max_val Then
        segment = (g - b) / C
        shift = 0 / 60
        If segment < 0 Then shift = 360 / 60
        hue = segment + shift
    ElseIf g = max_val Then
        segment = (b - r) / C
        shift = 120 / 60
        hue = segment + shift
    ElseIf b = max_val Then
        segment = (r - g) / C
        shift = 240 / 60
        hue = segment + shift
    End If
End If

hue = hue * 60
If max_val = 0 Then
    sat = 0
Else
    sat = C / max_val
End If
val_l = max_val
        
hsv(1) = hue
hsv(2) = sat
hsv(3) = val_l

RGBtoHSV = hsv
End Function

Function RGBtoHSL(rgb As Variant) As Variant
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'convert colours from RGB to HSL.
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Dim hsl(1 To 3) As Variant
red = rgb(1)
green = rgb(2)
blue = rgb(3)


r = red / 255
g = green / 255
b = blue / 255

max_val = Application.WorksheetFunction.Max(r, g, b)
min_val = Application.WorksheetFunction.Min(r, g, b)
C = max_val - min_val

If C = 0 Then
    hue = 0
Else
    If r = max_val Then
        hue = 60 * (0 + (g - b) / C)
    ElseIf g = max_val Then
        hue = 60 * (2 + (b - r) / C)
    ElseIf b = max_val Then
        hue = 60 * (4 + (r - g) / C)
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
Function HSLtoRGB(hsl As Variant) As Variant
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'convert colours from HSL to RGB.
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Dim rgb(1 To 3) As Variant
h = hsl(1)
s = hsl(2)
L = hsl(3)

C = (1 - Abs(2 * L - 1)) * s
hp = h / 60
mod_f = XLMod(hp, 2)
x = C * (1 - Abs((mod_f) - 1))


If h >= 0 And h < 60 Then
    r = C
    g = x
    b = 0
ElseIf h >= 60 And h < 120 Then
    r = x
    g = C
    b = 0
ElseIf h >= 120 And h < 180 Then
    r = 0
    g = C
    b = x
ElseIf h >= 180 And h < 240 Then
    r = 0
    g = x
    b = C
ElseIf h >= 240 And h < 300 Then
    r = x
    g = 0
    b = C
ElseIf h >= 300 And h <= 360 Then
    r = C
    g = 0
    b = x
Else
    r = 0
    g = 0
    b = 0
End If
M = L - C / 2

rgb(1) = (r + M) * 255
rgb(2) = (g + M) * 255
rgb(3) = (b + M) * 255
HSLtoRGB = rgb
End Function
Function RGBtoLight(rgb As Variant) As Integer
red = rgb(1)
green = rgb(2)
blue = rgb(3)

RGBtoLight = 0.2126 * red + 0.7152 * green + 0.0722 * blue
End Function

Sub BubbleSort(Arr)
  Dim strTemp As String
  Dim i As Long
  Dim J As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(Arr)
  lngMax = UBound(Arr)
  For i = lngMin To lngMax - 1
    For J = i + 1 To lngMax
      If GetNumber(Arr(i)) > GetNumber(Arr(J)) Then
        strTemp = Arr(i)
        Arr(i) = Arr(J)
        Arr(J) = strTemp
      End If
    Next J
  Next i
End Sub

Sub ColorOrderingHSL()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'sort colours based on HSL values
'+++++++++++++++++++++++++++++++++++++++++++++++++++
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
Dim Arr(1 To 2) As Variant
Dim hsl(1 To 3) As Variant

For i = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        'Debug.Print red; green; blue
        
        
        v(1) = red
        v(2) = green
        v(3) = blue
        
        colHSL = RGBtoHSL(v)
        h1 = Format((colHSL(1) / 360), "#0.000")
        s2 = Format((colHSL(2)), "#0.000")
        L3 = Format((colHSL(3)), "#0.000")
        
        
        colSortN = Round(5 * colHSL(3) + colHSL(2) * 2 + colHSL(1), 4)
        
        string_col = "H" & h1 & "S" & s2 & "L" & L3
        colors(i, J) = string_col
        
        colMat(i, J) = colSortN

    Next J
Next i


For i = 1 To 40
    slice = Application.Transpose(Application.WorksheetFunction.Index(colMat, 0, i))
    slice2 = Application.Transpose(Application.WorksheetFunction.Index(colors, 0, i))
    Call sort2(slice(), slice2())
    For J = 1 To 40

        orderedMat(J, i) = slice(J)
        orderedMatCol(J, i) = slice2(J)
    Next J
Next i

For i = 1 To 40
    slice = Application.WorksheetFunction.Index(orderedMat, i, 0)
    slice2 = Application.WorksheetFunction.Index(orderedMatCol, i, 0)
    Call sort2(slice(), slice2())
    For J = 1 To 40

        orderedMatFinal(J, i) = slice2(J)
    Next J
Next i

For i = 1 To 40
    For J = 1 To 40

        clr = orderedMatFinal(i, J)
        h = CDbl(Right(Left(clr, 6), 5))
        s = CDbl(Left(Right(clr, 11), 5))
        L = CDbl(Right(clr, 5))
        
        hsl(1) = h * 360
        Debug.Print (h)
        hsl(2) = s
        hsl(3) = L
        rgb_cols = HSLtoRGB(hsl)
'        Debug.Print rgb_cols(1); rgb_cols(2); rgb_cols(3)
        cells(27 + i, 12 + J).Interior.Color = rgb(rgb_cols(1), rgb_cols(2), rgb_cols(3))

    Next J
Next i
'ws_pic.Range("M28:AZ67") = orderedMatFinal
'ws_t.Range(ws_t.cells(1, 1), ws_t.cells(39, 39)) = orderedMatFinal
'Sheets("Temp").Delete
Application.DisplayAlerts = True
End Sub

Function getSortCol(s As Double) As Double

L = Len(s)
s = Left(s, Find("H", s) - 1)

getSortCol = s
End Function


Sub sort2(key() As Variant, other() As Variant)
Dim i As Long, J As Long, Low As Long
Dim Hi As Long, Temp As Variant
    Low = LBound(key)
    Hi = UBound(key)

    J = (Hi - Low + 1) \ 2
    Do While J > 0
        For i = Low To Hi - J
          If key(i) > key(i + J) Then
            Temp = key(i)
            key(i) = key(i + J)
            key(i + J) = Temp
            Temp = other(i)
            other(i) = other(i + J)
            other(i + J) = Temp
          End If
          
        Next i
        For i = Hi - J To Low Step -1
          If key(i) > key(i + J) Then
            Temp = key(i)
            key(i) = key(i + J)
            key(i + J) = Temp
            Temp = other(i)
            other(i) = other(i + J)
            other(i + J) = Temp
          End If
        Next i
        J = J \ 2
    Loop
End Sub

Sub testingHSLtoRGB()
Dim rgb(1 To 3) As Variant
h = 275.2174
s = 0.380165
L = 0.47451

C = (1 - Abs(2 * L - 1)) * s


hp = h / 60
mod_f = XLMod(hp, 2)
'mod_f = hp Mod 2
x = C * (1 - Abs((mod_f) - 1))

Debug.Print "h: "; h; "x: "; x

If h >= 0 And h < 60 Then
    r = C
    g = x
    b = 0
ElseIf h >= 60 And h < 120 Then
    r = x
    g = C
    b = 0
ElseIf h >= 120 And h < 180 Then
    r = 0
    g = C
    b = x
ElseIf h >= 180 And h < 240 Then
    r = 0
    g = x
    b = C
ElseIf h >= 240 And h < 300 Then
    r = x
    g = 0
    b = C
ElseIf h >= 300 And h <= 360 Then
    r = C
    g = 0
    b = x
Else
    r = 0
    g = 0
    b = 0
End If
M = L - C / 2

rgb(1) = (r + M) * 255
rgb(2) = (g + M) * 255
rgb(3) = (b + M) * 255
Debug.Print "red: "; rgb(1); "green: "; rgb(2); "blue: "; rgb(3)
End Sub
Sub RGBChart3d()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'Not working
'+++++++++++++++++++++++++++++++++++++++++++++++++++

Dim loc As Range
Dim cht As ChartObject

Set ws_pic = Sheets("SCARICHI BLOOMBERG")
Set Location = ws_pic.Range("N72:AZ95")
Application.DisplayAlerts = False

Dim red_array(1 To 10, 1 To 1) As Variant
Dim green_array(1 To 10, 1 To 1) As Variant
Dim blue_array(1 To 10, 1 To 1) As Variant

counter = 0
For i = 1 To 10
    For J = 1 To 10
        counter = counter + 1
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
              
        red_array(i, 1) = red
        green_array(i, 1) = green
        blue_array(i, 1) = blue

     Next J
Next i

'With Location
'    Set cht = ws_pic.ChartObjects.Add(.Left, .Top, .Width, .Height)
'    cht.Name = "3D Chart"
'End With
'
'With cht.Chart
'    .ChartType = xlColumnClustered
'
'    For k = 1 To UBound(red_array)
'        With .SeriesCollection.NewSeries
'            .XValues = blue_array
'            .Values = Application.WorksheetFunction.Index(red_array, 0, k)
'            .Name = green_array(k)
'        End With
'    Next k
'    .ChartType = xlSurface
'End With


ws_pic.Range("C69:C78") = red_array
ws_pic.Range("D69:D78") = green_array
ws_pic.Range("E69:E78") = blue_array
Application.DisplayAlerts = True
End Sub
Function XLMod(a, b)
    XLMod = a - b * Int(a / b)
End Function

Sub RandomColBand()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'create colour bands
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'For Each obj In ActiveSheet.OLEObjects
'    If obj.progID = "Forms.CheckBox.1" Then
'    Debug.Print obj.Name
'    End If
'Next obj

If ActiveSheet.CheckBox1.Value Then a = 1 Else a = 0
If ActiveSheet.CheckBox2.Value Then b = 1 Else b = 0
If ActiveSheet.CheckBox3.Value Then C = 1 Else C = 0

If ActiveSheet.ToggleButton1.Value Then b_w = 255 Else b_w = 0
 
sum = a + b + C
If sum = 0 Then
    ActiveSheet.CheckBox4.Value = True
    sum = 3
End If

If sum = 3 Then
    
    For Each cell In Range("M28:Z67")
        cell.Interior.Color = rgb(WorksheetFunction.RandBetween(0, 255), b_w, b_w)
    Next cell
    
    For Each cell In Range("AA28:AM67")
        cell.Interior.Color = rgb(b_w, WorksheetFunction.RandBetween(0, 255), b_w)
    Next cell
    
    For Each cell In Range("AN28:AZ67")
        cell.Interior.Color = rgb(b_w, b_w, WorksheetFunction.RandBetween(0, 255))
    Next cell
    
End If

End Sub

Sub circleSquare()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'create concentric squares
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Dim cell As Range
If ActiveSheet.ToggleButton1.Value Then b_w = 255 Else b_w = 0
w = 40

b = Int(ActiveSheet.TextBox2.Value)
If IsNumeric(b) = False Or b > 11 Then b = 1

Count = b

While w > 0
    

    If Count > 20 Then GoTo nit
    If b = 1 Then
        For Each cell In Range(cells(27 + Count, 12 + Count), cells(68 - Count, 53 - Count))
        
    
            If idx = 0 Then
                cell.Interior.Color = rgb(Application.WorksheetFunction.RandBetween(0, 255), b_w, b_w)
            ElseIf idx = 1 Then
               cell.Interior.Color = rgb(b_w, Application.WorksheetFunction.RandBetween(0, 255), b_w)
            Else
               cell.Interior.Color = rgb(b_w, b_w, Application.WorksheetFunction.RandBetween(0, 255))
            End If
    
            
            Next cell
            Count = Count + b
    '        Debug.Print Count
            w = w - b
            
            idx = idx + 1
            
            If idx = 3 Then idx = 0

    Else
        For Each cell In Range(cells(26 + Count, 11 + Count), cells(69 - Count, 54 - Count))
        
    
            If idx = 0 Then
                cell.Interior.Color = rgb(Application.WorksheetFunction.RandBetween(0, 255), b_w, b_w)
            ElseIf idx = 1 Then
               cell.Interior.Color = rgb(b_w, Application.WorksheetFunction.RandBetween(0, 255), b_w)
            Else
               cell.Interior.Color = rgb(b_w, b_w, Application.WorksheetFunction.RandBetween(0, 255))
            End If
    
            
            Next cell
            Count = Count + b
    '        Debug.Print Count
            w = w - b
            
            idx = idx + 1
            
            If idx = 3 Then idx = 0
        End If
Wend
nit:
For Each cell In Range("M28:AZ67")

    If Filled(cell) = "" Then
        cell.Interior.Color = rgb(Application.WorksheetFunction.RandBetween(0, 255), Application.WorksheetFunction.RandBetween(0, 255), Application.WorksheetFunction.RandBetween(0, 255))
    End If
Next cell

End Sub

Function Filled(MyCell As Range)
Filled = IIf(MyCell.Interior.ColorIndex = xlNone, "", 1)
End Function


Sub ShufflePix()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'randomly shuffles pixels
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Dim colMat(1 To 40, 1 To 40) As Variant
Dim orderedMat(1 To 40, 1 To 40) As Variant
Dim orderedMatFinal(1 To 40, 1 To 40) As Variant
Dim slice()

For i = 1 To 40
    For J = 1 To 40
        colMat(i, J) = cells(27 + i, 12 + J).DisplayFormat.Interior.Color
    Next J
Next i



For i = 1 To 40
    slice = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.Index(colMat, 0, i))
    ShuffleArrayInPlace slice
    For J = 1 To 40

        orderedMat(J, i) = slice(J)
    Next J
Next i

For i = 1 To 40
    slice = Application.WorksheetFunction.Index(orderedMat, i, 0)
    
    ShuffleArrayInPlace slice
    For J = 1 To 40
        orderedMatFinal(J, i) = slice(J)
    Next J
Next i


For i = 1 To 40
    For J = 1 To 40
        
        clr = orderedMatFinal(i, J)
        red = clr And 255
        green = clr \ 256 And 255
        blue = clr \ 256 ^ 2 And 255

        cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)

    Next J
Next i



End Sub

Sub square_blur()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'blurring filters
'+++++++++++++++++++++++++++++++++++++++++++++++++++

s_side = Int(ActiveSheet.TextBox3.Value)
If IsNumeric(s_side) = False Or s_side < 2 Then s_side = 2

's_side = 3
b_side = 40

n_square = b_side ^ 2 - 2 * b_side * s_side + 2 * b_side - 2 * s_side + s_side ^ 2 + 1

'Debug.Print n_square
N = 1
M = 1
While M <= b_side - s_side + 1
    While N <= b_side - s_side + 1
    
        Set square = Range(cells(27 + N, 12 + M), cells(27 + N + s_side - 1, 12 + M + s_side - 1))
            red_c = 0
            green_c = 0
            blue_c = 0
            For Each cell In square
                C = cell.Interior.Color
                red_c = C Mod 256 + red_c
                green_c = C \ 256 Mod 256 + green_c
               blue_c = C \ 65536 Mod 256 + blue_c
            Next cell
            red_f = red_c / (s_side ^ 2)
            green_f = green_c / (s_side ^ 2)
            blue_f = blue_c / (s_side ^ 2)
            
            square.Interior.Color = rgb(red_f, green_f, blue_f)
            
            

        
        Debug.Print N
        N = N + 1
        
    Wend
    N = 1
    M = M + 1
Wend

End Sub

Sub posterize()
Set ws_pic = Sheets("SCARICHI BLOOMBERG")
For i = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        If red < 128 Then red = 0 Else red = 255
        If green < 128 Then green = 0 Else green = 255
        If blue < 128 Then blue = 0 Else blue = 255
        
        
                cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)


       
    Next J
Next i



End Sub

Sub inverse_posterize()
Set ws_pic = Sheets("SCARICHI BLOOMBERG")
For i = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        If red < 128 Then red = 255 Else red = 0
        If green < 128 Then green = 255 Else green = 0
        If blue < 128 Then blue = 255 Else blue = 0
        
        
                cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)


       
    Next J
Next i



End Sub

Sub DecreaseColor()
Set ws_pic = Sheets("SCARICHI BLOOMBERG")
For i = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        If ActiveSheet.CheckBox2.Value = True Then red = red * 0.9
        If ActiveSheet.CheckBox1.Value = True Then green = green * 0.9
        If ActiveSheet.CheckBox3.Value = True Then blue = blue * 0.9
        
        
        
        cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)


       
    Next J
Next i



End Sub

Sub IncreaseColor()
Set ws_pic = Sheets("SCARICHI BLOOMBERG")
For i = 1 To 40
    For J = 1 To 40
        col = ws_pic.cells(27 + i, 12 + J).DisplayFormat.Interior.Color
        red = col And 255
        green = col \ 256 And 255
        blue = col \ 256 ^ 2 And 255
        
        If ActiveSheet.CheckBox2.Value = True Then red = red * 1.1
        If ActiveSheet.CheckBox1.Value = True Then green = green * 1.1
        If ActiveSheet.CheckBox3.Value = True Then blue = blue * 1.1
        
        If red > 255 Then red = 255
        If green > 255 Then green = 255
        If blue > 255 Then blue = 255
        
        cells(27 + i, 12 + J).Interior.Color = rgb(red, green, blue)


       
    Next J
Next i



End Sub
Sub Interpol_blur()
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'blurring filters
'+++++++++++++++++++++++++++++++++++++++++++++++++++

s_side = 3
b_side = 40

n_square = b_side ^ 2 - 2 * b_side * s_side + 2 * b_side - 2 * s_side + s_side ^ 2 + 1

'Debug.Print n_square
N = 1
M = 1
While M <= b_side - s_side + 1
    While N <= b_side - s_side + 1
    
        'Set square = Range(cells(27 + N, 12 + M), cells(27 + N + s_side - 1, 12 + M + s_side - 1))
        
            Gdelta_v = Abs((cells(27 + N, 13 + M).Interior.Color \ 256 Mod 256) - (cells(29 + N, 13 + M).Interior.Color \ 256 Mod 256))
            Gdelta_h = Abs((cells(28 + N, 12 + M).Interior.Color \ 256 Mod 256) - (cells(28 + N, 14 + M).Interior.Color \ 256 Mod 256))
            
            If Gdelta_h = Gdelta_v Then
                green = ((cells(27 + N, 13 + M).Interior.Color \ 256 Mod 256) + (cells(29 + N, 13 + M).Interior.Color \ 256 Mod 256) + (cells(28 + N, 12 + M).Interior.Color \ 256 Mod 256) + (cells(28 + N, 14 + M).Interior.Color \ 256 Mod 256)) / 2
            ElseIf Gdelta_h > Gdelta_v Then
                green = ((cells(27 + N, 13 + M).Interior.Color \ 256 Mod 256) + (cells(29 + N, 13 + M).Interior.Color \ 256 Mod 256)) / 2
            Else
                green = ((cells(28 + N, 12 + M).Interior.Color \ 256 Mod 256) + (cells(28 + N, 14 + M).Interior.Color \ 256 Mod 256)) / 2
            End If
                
            Rdelta_v = Abs((cells(27 + N, 13 + M).Interior.Color Mod 256) - (cells(29 + N, 13 + M).Interior.Color Mod 256))
            Rdelta_h = Abs((cells(28 + N, 12 + M).Interior.Color Mod 256) - (cells(28 + N, 14 + M).Interior.Color Mod 256))
            
            If Rdelta_h = Rdelta_v Then
                red = ((cells(27 + N, 13 + M).Interior.Color Mod 256) + (cells(29 + N, 13 + M).Interior.Color Mod 256) + (cells(28 + N, 12 + M).Interior.Color Mod 256) + (cells(28 + N, 14 + M).Interior.Color Mod 256)) / 2
            ElseIf Rdelta_h > Rdelta_v Then
                red = ((cells(27 + N, 13 + M).Interior.Color Mod 256) + (cells(29 + N, 13 + M).Interior.Color Mod 256)) / 2
            Else
                red = ((cells(28 + N, 12 + M).Interior.Color Mod 256) + (cells(28 + N, 14 + M).Interior.Color Mod 256)) / 2
            End If
                
            Bdelta_v = Abs((cells(27 + N, 13 + M).Interior.Color \ 65536 Mod 256) - (cells(29 + N, 13 + M).Interior.Color \ 65536 Mod 256))
            Bdelta_h = Abs((cells(28 + N, 12 + M).Interior.Color \ 65536 Mod 256) - (cells(28 + N, 14 + M).Interior.Color \ 65536 Mod 256))
            
            If Bdelta_h = Bdelta_v Then
                blue = ((cells(27 + N, 13 + M).Interior.Color \ 65536 Mod 256) + (cells(29 + N, 13 + M).Interior.Color \ 65536 Mod 256) + (cells(28 + N, 12 + M).Interior.Color \ 65536 Mod 256) + (cells(28 + N, 14 + M).Interior.Color \ 65536 Mod 256)) / 2
            ElseIf Bdelta_h > Bdelta_v Then
                blue = ((cells(27 + N, 13 + M).Interior.Color \ 65536 Mod 256) + (cells(29 + N, 13 + M).Interior.Color \ 65536 Mod 256)) / 2
            Else
                blue = ((cells(28 + N, 12 + M).Interior.Color \ 65536 Mod 256) + (cells(28 + N, 14 + M).Interior.Color \ 65536 Mod 256)) / 2
                
            cells(28 + N, 13 + M).Interior.Color = rgb(red, green, blue)
            End If
'            red_c = 0
'            green_c = 0
'            blue_c = 0
'            For Each cell In square
'                C = cell.Interior.Color
'                red_c = C Mod 256 + red_c
'                green_c = C \ 256 Mod 256 + green_c
'               blue_c = C \ 65536 Mod 256 + blue_c
'            Next cell
'            red_f = red_c / (s_side ^ 2)
'            green_f = green_c / (s_side ^ 2)
'            blue_f = blue_c / (s_side ^ 2)
'
'            square.Interior.Color = rgb(red_f, green_f, blue_f)
                   
        N = N + 1
        
    Wend
    N = 1
    M = M + 1
Wend

End Sub

Private Sub Reverse_Array_2d(ByRef Ary As Variant, Optional Header_Rows As Integer = 0)

 Dim Dimension_Y As Integer     ' Rows (height)
 Dim Y_first As Long
 Dim Y_last As Long
 Dim Y_last_plus_Y_first As Long
 Dim Y_next As Long

 Dimension_Y = 1
 Y_first = LBound(Ary, Dimension_Y) + Header_Rows
 Y_last = UBound(Ary, Dimension_Y)
 Y_last_plus_Y_first = Y_last + Y_first

 Dim Dimension_X As Integer      ' Columns (width)
 Dim X_first As Long
 Dim X_last As Long

 Dimension_X = 2
 X_first = LBound(Ary, Dimension_X)
 X_last = UBound(Ary, Dimension_X)

 ReDim tmp(X_first To X_last) As Variant

 For Y = Y_first To Y_last_plus_Y_first / 2
    Y_next = Y_last_plus_Y_first - Y
    For x = X_first To X_last
        tmp(x) = Ary(Y_next, x)
        Ary(Y_next, x) = Ary(Y, x)
        Ary(Y, x) = tmp(x)
    Next
 Next

End Sub
Sub flip_pic()
Dim colMat(1 To 40, 1 To 40) As Variant
Dim orderedMat(1 To 40, 1 To 40) As Variant
Dim orderedMatFinal(1 To 40, 1 To 40) As Variant
Dim slice()

For i = 1 To 40
    For J = 1 To 40
        colMat(i, J) = cells(27 + i, 12 + J).DisplayFormat.Interior.Color
    Next J
Next i

Call Reverse_Array_2d(colMat, CInt(0))

For i = 1 To 40
    For J = 1 To 40
        cells(27 + i, 12 + J).Interior.Color = colMat(i, J)
    Next J
Next i


End Sub

