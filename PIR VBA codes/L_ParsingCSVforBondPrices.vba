Sub parserCSV()
Dim ws As Worksheet
Dim date_file_to_parse As String

'Set ws = Sheets("Foglio2")

date_file_to_parse = Sheets("STORICO PREZZI_FI").Range("F2")

Application.Calculation = xlCalculationManual

lstRow = Sheets("STORICO PREZZI_FI").Range("C1")

'dat3e handling to retrive file

strDate = date_file_to_parse
date_file = date_format(date_file_to_parse)
dt_day = Format(date_file, "m")
dt_month = Format(date_file, "mmmm")
dt_month = Application.Proper(dt_month)
dt_path = dt_day & "-" & dt_month
dt_year = Format(date_file, "yyyy")

    Close #1
    s = "Y:\Mobiliare\08 Finint Economia Reale Italia\03_Back Office\04_Prezzi\obbligazionario\" & _
            dt_year & "\" & dt_path & "\prezzi_" & strDate & ".csv"
Debug.Print s
    If checkFileExistance(s) Then
    Else
        On Error GoTo retError
        s = "Y:\Mobiliare\08 Finint Economia Reale Italia\03_Back Office\04_Prezzi\obbligazionario\" & _
            dt_year & "\prezzi_" & strDate & ".csv"
    End If
    
    Open s For Input As #1

   J = 1
   Do While Not EOF(1)
      Line Input #1, TextLine
      Ary = Split(TextLine, ";")
      i = 1
      For Each a In Ary
        If i = 3 Then
            a = str(a)
            a = Replace(a, ".", ",")
            a = CDbl(a)
        End If
        cells(J + lstRow, i).Value = a
      i = i + 1

      
      Next a
    J = J + 1
   Loop

   Close #1
   

 
Application.Calculation = xlCalculationAutomatic

lrafter = Sheets("STORICO PREZZI_FI").Range("C1")
With Sheets("STORICO PREZZI_FI").Range(cells(lrafter, 1), cells(lrafter, 6)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
        
End With

Done:
   Exit Sub
   
retError:
 MsgBox "File Not Found"
 Application.Calculation = xlCalculationAutomatic
End Sub

Private Function checkFileExistance(ByVal strFolderPath As String) As Boolean
 
checkFileExistance = Dir(strFolderPath, vbDirectory) <> vbNullString

End Function

Function date_format(date_used As String)

date_used = Left(date_used, 2) & "/" & Mid(date_used, 3, 2) & "/" & Right(date_used, 4)

date_format = Format(date_used, "DD/MMM/YYYY")
End Function