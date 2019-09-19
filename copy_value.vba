Sub Sample()
    Dim thisWb As Workbook, wbTemp As Workbook
    Dim ws As Worksheet

    On Error GoTo Whoa

    Application.DisplayAlerts = False

    Set thisWb = ThisWorkbook
    Set wbTemp = Workbooks.Add

    On Error Resume Next
    For Each ws In wbTemp.Worksheets
        ws.Delete
    Next
    On Error GoTo 0

    For Each ws In thisWb.Sheets
        ws.Copy After:=wbTemp.Sheets(1)
    Next

    wbTemp.Sheets(1).Delete
    wbTemp.SaveAs "C:\Users\rossetti01\Documents\prova copy value\Blah Blah.xls", xlWorkbookNormal
    wbTemp.Close
LetsContinue:
    Application.DisplayAlerts = True
    Exit Sub
Whoa:
    MsgBox Err.Description
    Resume LetsContinue
End Sub
