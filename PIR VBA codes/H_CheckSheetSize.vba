'CODE FOR PRINTING IN WINDOW THE SIZE OF EACH OF THE SHEETS IN THE WORKBOOK FILE
'
'



Sub GetSheetSizes()
' ZVI:2012-05-18 Excel VBA File Size by Worksheet in File
' CAR:2014-10-07 Enhanced to take hidden and very hidden sheets into account
  
  Dim a() As Variant
  Dim Bytes As Double
  Dim i As Long
  Dim fileNameTmp As String
  Dim wb As Workbook
  Dim visState As Integer
  
  Set wb = ActiveWorkbook
  ReDim a(0 To wb.Sheets.Count, 1 To 2)
  
  ' Turn off screen updating
  Application.ScreenUpdating = False
  On Error GoTo exit_
  
  ' Put names into a(,1) and sizes into a(,2)
  With CreateObject("Scripting.FileSystemObject")
    ' Build the temporary file name
    Err.Clear
    fileNameTmp = .GetSpecialFolder(2) & "\" & wb.Name & ".TMP"
    ' Put workbook's name and size into a(0,)
    a(0, 1) = wb.Name
    a(0, 2) = .GetFile(wb.FullName).Size
    ' Put each sheet name and its size into a(i,)
    For i = 1 To wb.Sheets.Count
      visState = wb.Sheets(i).Visible
      wb.Sheets(i).Visible = -1 ' Show sheet long enough to copy it
      DoEvents
      wb.Sheets(i).Copy
      
      ActiveWorkbook.SaveCopyAs fileNameTmp
      
      wb.Sheets(i).Visible = visState
      a(i, 1) = wb.Sheets(i).Name
      a(i, 2) = .GetFile(fileNameTmp).Size
      Bytes = Bytes + a(i, 2)
      ActiveWorkbook.Close False
    Next
    Kill fileNameTmp
  End With
  
  ' Show workbook's name & size
  Debug.Print a(0, 1), Format(a(0, 2), "#,##0") & " Bytes"
  
  ' Show workbook object's  size
  Debug.Print "Wb Object", Format(a(0, 2) - Bytes, "#,##0") & " Bytes"
  
  ' Show each sheet name and its size
  For i = 1 To UBound(a)
    Debug.Print a(i, 1), Format(a(i, 2), "#,##0") & " Bytes"
  Next
  
exit_:
  
  ' Restore screen updating
  Application.ScreenUpdating = True
  
  ' Show the reason of error if happened
  If Err Then MsgBox Err.Description, vbCritical, "Error"


End Sub