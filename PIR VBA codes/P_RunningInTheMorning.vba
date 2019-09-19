'Sub CallingRoutine()
'Application.OnTime TimeSerial(6, 15, 0), "CallingRoutine"
'Workbooks("Portafoglio PIR.xlsm").Activate
'ThisWorkbook.Sheets("Tables").Activate
'Call WeDontNeedNoEduMail
'
'DoEvents
'    Application.Wait (Now + TimeValue("0:01:00"))
'Call Open_Outlook
'End Sub
'
'
'Sub ref()
'Application.Run "RefreshEntireWorkbook"
'Application.Run "RefreshAllStaticData"
'
'End Sub
'
'Sub cRef()
'
'Call ref
'Application.OnTime TimeSerial(6, 0, 0), "cRef"
'
'End Sub