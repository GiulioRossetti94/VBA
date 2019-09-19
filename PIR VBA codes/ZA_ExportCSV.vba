Sub Sheets_to_csv()

Application.Calculation = xlCalculationAutomatic

file_name = ActiveSheet.Name
Application.DisplayAlerts = False
strFullname = "C:\Users\bloomberg03\Desktop\PythonScript\CSV\" & file_name
ActiveWorkbook.ActiveSheet.Copy
ActiveWorkbook.SaveAs FileName:=strFullname, FileFormat:=xlCSV, CreateBackup:=False
ActiveWorkbook.ActiveSheet.UsedRange.NumberFormat = "General"
ActiveWorkbook.Close

Application.DisplayAlerts = True
End Sub