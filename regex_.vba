Option Base 1
Dim R As Integer
Dim a As Integer
Dim b As Integer
Dim i As Integer
Dim j As Integer

'========================================================================================================================
Sub final()
Call verPrezzi
Call ID
Call Isin
Call name_space
Call del_blank_cells

Dim ar As Variant
Dim rng As Range
Sheets("Ric").Activate
lr = Cells(1, 4).End(xlDown).Row
Set rng = Range(Cells(1, 4), Cells(lr, 4))
ar = rng.Value

For lambda = LBound(ar, 1) To UBound(ar, 1)
    If ar(lambda, 1) > 1 Then ar(lambda, 1) = ar(lambda, 1) / 1000 Else ar(lambda, 1) = CDbl(ar(lambda, 1))
Next
rng.Value = ar

End Sub
'========================================================================================================================
Sub verPrezzi()
Dim rng As Range
Dim txt As String
Dim i As Integer
Dim FullName As Variant
Dim x As String
Dim c As Integer
'Range("A1:A35").EntireRow.Delete

c = 0
ActiveSheet.Name = "Txt"
Sheets.Add.Name = "Copia"
Sheets.Add.Name = "Ric"
Sheets("Txt").Activate

Set rng = Range("A:A")

With CreateObject("vbscript.regexp")
  .Pattern = "\s{2,}"
  .Global = True
  For Each cell In Selection.SpecialCells(xlCellTypeConstants)
    cell.Value = .Replace(cell.Value, " ")
  Next cell
End With

For Each cell In ActiveSheet.Range(Cells(1, 1), Cells(ActiveSheet.UsedRange.Count, 1))
     txt = cell.Value

     FullName = Split(txt, " ")

     For i = 0 To UBound(FullName)

         cell.Offset(0, i + 2).Value = FullName(i)
     If c < i + 1 Then c = i + 3
     Next i

Next cell

e = Cells(1, 1).End(xlDown).Row

Range(Cells(1, 2), Cells(e, c)).Cut
Sheets("Copia").Activate
ActiveSheet.Paste
Application.CutCopyMode = False
End Sub

'========================================================================================================================
Sub ID()
'primo loop su  cod int
Dim R As Integer

R = 142
a = 2
b = 2
i = 9
j = 1
For j = 1 To R

check = Sheets("Copia").Cells(j, a)
If IsNumeric(check) Then Sheets("Ric").Cells(j, a) = check

Next

End Sub
'========================================================================================================================
Sub Isin()

R = 142
a = 3
b = 2
i = 9
j = 1
For j = 1 To R

check = Sheets("Ric").Cells(j, a - 1)
If Not IsEmpty(check) Then Sheets("Ric").Cells(j, a) = Sheets("Copia").Cells(j, a)

Next
End Sub

Sub name_space()
Dim text As String
Dim rng As Range
Dim arr As Variant
R = 142
a = 4
b = 2
i = 9
j = 1
Sheets("Copia").Activate

For j = 1 To R

    
     Range(Cells(j, 3), Cells(j, 11)).Select
        For Each k In Selection
    check = Sheets("Ric").Cells(j, a - 1)
       text = k.text
        If text Like "EUR" Or text Like "OEUR" Or text Like "UCITEUR" Then
       cD = k.Column
If Not IsEmpty(check) Then Sheets("Ric").Cells(j, a) = Sheets("Copia").Cells(j, cD + 3)
        End If
        
 Next


Next j


End Sub
'========================================================================================================================
Sub del_blank_cells()

Dim kappa As Range
Dim r1 As Integer
Dim phi As Integer

Sheets("Ric").Activate
Set kappa = Range("B1:D200")
nR = kappa.Rows.Count

For phi = nR To 1 Step (-1)
   If WorksheetFunction.CountA(kappa.Rows(phi)) = 0 Then kappa.Rows(phi).Delete
  Next
 

End Sub
