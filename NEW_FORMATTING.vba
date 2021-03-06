Sub call_formatting()
Application.StatusBar = "NINO NINO NINO NINO"
date_report = Sheets("Pir performance").Cells(1, 1).End(xlDown)
Sheets("First Page").Range("F26") = date_report
Sheets("First Page").Range("F27") = Format(Application.WorksheetFunction.WorkDay(Application.WorksheetFunction.EoMonth(Now(), -1) + 1, -1), "yyyymmdd")
Sheets("First Page").Range("F28") = 20181228

Call format_perf_sheet
Call formatting_portfolio_sheets
Call formatting_summary_new
Call formatting_tables_asset_class
Call format_subs_sheet
Call VAR_format
Application.StatusBar = False
End Sub




Sub format_perf_sheet()
Dim sheet_array() As Variant
Dim ws As Worksheet

Set ws = Worksheets("Performance sheet")
'FORMATTING PERFORMANCE TABLES
For Each i In Array(3, 18, 33)

    
    
    
    n_year = ws.Range(ws.Cells(4, i), ws.Cells(4, i).End(xlDown)).Count
    If n_year < 1 Then n_year = 1
    
    Set rng_all = ws.Range(ws.Cells(4, i), ws.Cells(n_year + 3, i + 13))
    Set head = ws.Range(ws.Cells(4, i), ws.Cells(4, i + 13))
    Set body = ws.Range(ws.Cells(5, i + 1), ws.Cells(n_year + 5, i + 13))
    Set body_index = ws.Range(ws.Cells(5, i), ws.Cells(n_year + 5 - 2, i + 13))
   
    With rng_all
        .ClearFormats
        .HorizontalAlignment = xlHAlignCenter
        .FormatConditions.Delete
    End With
    
    With head
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
            
    End With
    
    ws.Range(ws.Cells(4, 3), ws.Cells(n_year + 3, 3)).Font.Bold = True
    ws.Columns("C:AT").ColumnWidth = 8.5
    
    With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="n.d.")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 176, 80)
        .Font.Bold = True
    End With
    
    With body.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With
    
    With body_index
        .Borders(xlEdgeLeft).LineStyle = xlDot
        .Borders(xlEdgeRight).LineStyle = xlDot
        .Borders(xlEdgeTop).LineStyle = xlDouble
        .Borders(xlEdgeBottom).LineStyle = xlDouble
        .Borders(xlInsideVertical).LineStyle = xlDot
        .Borders(xlInsideHorizontal).LineStyle = xlDot
    End With
    body.NumberFormat = "0.00%"


Next i

'FORMATTING ROLLING ANALYSIS
Set roll_head = ws.Range(ws.Cells(16, 3), ws.Cells(16, 6))
Set roll_all = ws.Range(ws.Cells(16, 3), ws.Cells(32, 6))
Set roll_body = ws.Range(ws.Cells(17, 3), ws.Cells(32, 6))

With roll_all
    .ClearFormats
    .HorizontalAlignment = xlHAlignCenter
    .FormatConditions.Delete
End With
With roll_head
    .Font.Bold = True
    .Font.Color = RGB(255, 255, 255)
    .Interior.Color = RGB(0, 32, 96)
End With
With roll_body
    .Borders(xlEdgeLeft).LineStyle = xlDot
    .Borders(xlEdgeRight).LineStyle = xlDot
    .Borders(xlEdgeTop).LineStyle = xlDouble
    .Borders(xlEdgeBottom).LineStyle = xlDouble
    .Borders(xlInsideVertical).LineStyle = xlDot
    .Borders(xlInsideHorizontal).LineStyle = xlDot
    .NumberFormat = "0.00%"
End With

With ws.Range(ws.Cells(16, 3), ws.Cells(32, 3))
    .Font.Bold = True
    .Font.Size = 8
    .HorizontalAlignment = xlHAlignLeft
End With

End Sub
Sub formatting_portfolio_sheets()
Dim ws_sh As Worksheet
Dim sheet_data() As Variant
Dim eqt_col_perc() As Variant
Dim eqt_col_bond() As Variant

sheet_data = Array("Equity port", "Bond port")

eqt_col_perc = Array(12, 13, 14, 15, 20)
bond_col_perc = Array(12, 13, 14, 15, 19)


For Each sh In sheet_data
    Set ws_sh = Worksheets(sh)
    
    n_row = ws_sh.Range(ws_sh.Cells(5, 3), ws_sh.Cells(5, 3).End(xlDown)).Count
    n_col = ws_sh.Range(ws_sh.Cells(5, 3), ws_sh.Cells(5, 3).End(xlToRight)).Count
    
    ws_sh.UsedRange.HorizontalAlignment = xlHAlignCenter
    ws_sh.Range("C:J").HorizontalAlignment = xlHAlignLeft
    
    With ws_sh.Range(ws_sh.Cells(5, 3), ws_sh.Cells(5, 3).End(xlToRight))
        .Interior.Color = RGB(154, 188, 230)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .EntireColumn.AutoFit
        .WrapText = True
    End With
    
     
    If InStr(sh, "Equity") > 0 Then
        For Each k In eqt_col_perc
           ws_sh.Columns(k).NumberFormat = "0.00%"
        Next k
        

        ws_sh.Columns(21).NumberFormat = "0.00"
        ws_sh.Columns(23).NumberFormat = "0.00"
        ws_sh.Columns(24).NumberFormat = "0.00"
        ws_sh.Columns(19).NumberFormat = "#,###0.00"
        ws_sh.Columns(11).NumberFormat = "#,###0.00"
        
    Else
        For Each k In bond_col_perc
           ws_sh.Columns(k).NumberFormat = "0.00%"
           
        Next k
        ws_sh.Columns(20).ColumnWidth = 10
        ws_sh.Columns(20).NumberFormat = "0.00"
        ws_sh.Columns(21).NumberFormat = "0.00"
        ws_sh.Columns(11).NumberFormat = "#,##0.00"

    End If

    
    
    
    ws_sh.Columns(12).ColumnWidth = 9
    ws_sh.Columns(16).NumberFormat = "0.00"
    ws_sh.Columns(17).NumberFormat = "0"
    
'conditional fromatting
    Set rng_exposure = ws_sh.Range(ws_sh.Cells(6, 12), ws_sh.Cells(4 + n_row, 12))
    Set rng_perf = ws_sh.Range(ws_sh.Cells(6, 13), ws_sh.Cells(4 + n_row, 15))
    Set rng_perf_canc = ws_sh.Range("L:O")
    
    With rng_perf_canc
        .FormatConditions.Delete
    End With
    
    With rng_exposure.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=5/100")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
        .Interior.Color = RGB(255, 255, 204)
        With .Borders
            .LineStyle = xlContinuous

            .Color = vbBlue
        End With
    End With
    
    With rng_perf.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="n.d.")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With rng_perf.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With rng_perf.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 176, 80)
        .Font.Bold = True
    End With
    
        With rng_perf.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With

Next sh

End Sub


Sub formatting_summary_new()

Dim ws_sh As Worksheet

Set ws_sh = Worksheets("Summary")

ncol_1 = Application.WorksheetFunction.CountA(ws_sh.Columns(3))
ncol_2 = Application.WorksheetFunction.CountA(ws_sh.Columns(7))
ncol_3 = Application.WorksheetFunction.CountA(ws_sh.Columns(11))

n = Application.WorksheetFunction.Max(ncol_1, ncol_2, ncol_3) + 3

Set urange = ws_sh.UsedRange

urange.Borders.LineStyle = xlNone

   
For i = 1 To n

    If ws_sh.Cells(2 + i, 3) = "" And ws_sh.Cells(4 + i, 3) <> "" And ws_sh.Cells(3 + i, 3) <> "" Then
        Set rng1 = ws_sh.Range(ws_sh.Cells(3 + i, 3), ws_sh.Cells(3 + i, 3).End(xlToRight))
        Set rng1_records = ws_sh.Range(ws_sh.Cells(4 + i, 3), ws_sh.Cells(4 + i, 3).End(xlDown))
        
        Set rng2 = ws_sh.Range(ws_sh.Cells(3 + i, 7), ws_sh.Cells(3 + i, 7).End(xlToRight))
        Set rng2_records = ws_sh.Range(ws_sh.Cells(4 + i, 7), ws_sh.Cells(4 + i, 7).End(xlDown))
        
        Set rng3 = ws_sh.Range(ws_sh.Cells(3 + i, 11), ws_sh.Cells(3 + i, 11).End(xlToRight))
        Set rng3_records = ws_sh.Range(ws_sh.Cells(4 + i, 11), ws_sh.Cells(4 + i, 11).End(xlDown))
    
        With rng1_records
            .ClearFormats
            .HorizontalAlignment = xlHAlignLeft
        End With
        
        With rng2_records
            .ClearFormats
            .HorizontalAlignment = xlHAlignLeft
        End With
        
        With rng3_records
         .ClearFormats
            .HorizontalAlignment = xlHAlignLeft
        End With
    
          
        With rng1
            If i < 8 Then
                .Interior.Color = RGB(0, 32, 96)
            Else
                .Interior.Color = RGB(154, 188, 230)
            End If
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
        End With
    
        With rng2
            If i < 8 Then
                .Interior.Color = RGB(0, 32, 96)
            Else
                .Interior.Color = RGB(154, 188, 230)
            End If
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
        End With
        
        With rng3
            If i < 8 Then
                .Interior.Color = RGB(0, 32, 96)
            Else
                .Interior.Color = RGB(154, 188, 230)
            End If
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
    
            .WrapText = True
        End With
        End If
Next i

ws_sh.Columns(3).Font.Bold = True
ws_sh.Columns(7).Font.Bold = True
ws_sh.Columns(11).Font.Bold = True

ws_sh.Columns(4).NumberFormat = "#,##0.00"
ws_sh.Columns(8).NumberFormat = "#,##0.00"
ws_sh.Columns(12).NumberFormat = "#,##0.00"

ws_sh.Columns(5).NumberFormat = "0.00%"
ws_sh.Columns(9).NumberFormat = "0.00%"
ws_sh.Columns(13).NumberFormat = "0.00%"

With ws_sh.Range(ws_sh.Cells(5, 4), ws_sh.Cells(5, 4).End(xlDown))
    .NumberFormat = "0.00%"
    .HorizontalAlignment = xlHAlignRight
End With

With ws_sh.Range(ws_sh.Cells(5, 8), ws_sh.Cells(5, 8).End(xlDown))
    .NumberFormat = "0.00%"
    .HorizontalAlignment = xlHAlignRight
End With

With ws_sh.Range(ws_sh.Cells(5, 12), ws_sh.Cells(5, 12).End(xlDown))
    .NumberFormat = "0.00%"
    .HorizontalAlignment = xlHAlignRight
End With

End Sub
Sub formatting_tables_asset_class()

sheet_data = Array("PIR Chart", "EQUITY Chart", "BOND Chart")

For Each sh In sheet_data
    Set ws_sh = Worksheets(sh)
    
    ws_sh.UsedRange.Borders.LineStyle = xlNone
    w = ws_sh.UsedRange.Rows.Count

    For i = 25 To w
        If ws_sh.Cells(2 + i, 3) = "" And ws_sh.Cells(4 + i, 3) <> "" And ws_sh.Cells(3 + i, 3) <> "" Then
            Set rng1 = ws_sh.Range(ws_sh.Cells(3 + i, 3), ws_sh.Cells(3 + i, 3).End(xlToRight))
            Set rng1_records = ws_sh.Range(ws_sh.Cells(4 + i, 3), ws_sh.Cells(4 + i, 3).End(xlDown))
            
            Set rng2 = ws_sh.Range(ws_sh.Cells(3 + i, 7), ws_sh.Cells(3 + i, 7).End(xlToRight))
            Set rng2_records = ws_sh.Range(ws_sh.Cells(4 + i, 7), ws_sh.Cells(4 + i, 7).End(xlDown))
            
        
            With rng1_records
                .ClearFormats
                .HorizontalAlignment = xlHAlignLeft
            End With
            
            With rng2_records
                .ClearFormats
                .HorizontalAlignment = xlHAlignLeft
            End With
              
              
            With rng1
                .Interior.Color = RGB(0, 32, 96)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .WrapText = True
            End With
        
            With rng2
                .Interior.Color = RGB(0, 32, 96)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .WrapText = True
            End With
            
        End If
    Next i

ws_sh.Columns(3).Font.Bold = True
ws_sh.Columns(7).Font.Bold = True
ws_sh.Columns(5).NumberFormat = "0.00%"
ws_sh.Columns(9).NumberFormat = "0.00%"

ws_sh.Range(ws_sh.Cells(27, 4), ws_sh.Cells(60, 4)).NumberFormat = "#,##0.00"
ws_sh.Range(ws_sh.Cells(27, 8), ws_sh.Cells(60, 8)).NumberFormat = "#,##0.00"

Set ret = ws_sh.Range(ws_sh.Cells(10, 4), ws_sh.Cells(12, 4))
If sh <> "PIR Chart" Then
    Set bw = ws_sh.Range(ws_sh.Cells(10, 9), ws_sh.Cells(16, 9))
    With bw
        .FormatConditions.Delete
    End With
    With bw.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With bw.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 176, 80)
        .Font.Bold = True
    End With
    
    With bw.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With

End If
    With ret
        .FormatConditions.Delete
    End With
    
    With ret.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0")
        .Font.Color = RGB(0, 32, 96)
        .Font.Bold = True
    End With
    
    With ret.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(0, 176, 80)
        .Font.Bold = True
    End With
    
    With ret.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With


Next sh
End Sub

Sub format_subs_sheet()
Set ws_sub = Worksheets("Subs")
ws_sub.UsedRange.Borders.LineStyle = xlNone


With ws_sub.Range("C5:D5")
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_sub.Range("F5:G5")
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_sub.Range(ws_sub.Cells(6, 3), ws_sub.Cells(6, 3).End(xlDown))
    .HorizontalAlignment = xlHAlignLeft
    .VerticalAlignment = xlVAlignCenter
End With

With ws_sub.Range(ws_sub.Cells(6, 6), ws_sub.Cells(6, 6).End(xlDown))
    .HorizontalAlignment = xlHAlignLeft
    .VerticalAlignment = xlVAlignCenter
End With

With ws_sub.Range(ws_sub.Cells(6, 7), ws_sub.Cells(6, 7).End(xlDown))
    .HorizontalAlignment = xlHAlignRight
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "#,##0.00"
End With

With ws_sub.Range(ws_sub.Cells(6, 4), ws_sub.Cells(6, 4).End(xlDown))
    .HorizontalAlignment = xlHAlignRight
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "#,##0.00"
End With

End Sub

Sub VAR_format()
Set ws_var = Worksheets("Var")
ws_var.UsedRange.Borders.LineStyle = xlNone
ws_var.Columns(3).ColumnWidth = 22

With ws_var.Range(ws_var.Cells(7, 4), ws_var.Cells(7, 4).End(xlToRight))
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_var.Range(ws_var.Cells(16, 4), ws_var.Cells(16, 4).End(xlToRight))
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_var.Range(ws_var.Cells(8, 4), ws_var.Cells(9, 4))
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0%"
End With

With ws_var.Range(ws_var.Cells(17, 4), ws_var.Cells(18, 4))
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0%"
End With

With ws_var.Range(ws_var.Cells(8, 5), ws_var.Cells(9, 9))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0.0000%"
End With

With ws_var.Range(ws_var.Cells(17, 5), ws_var.Cells(18, 9))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0.0000%"
End With
'

With ws_var.Range(ws_var.Cells(25, 4), ws_var.Cells(25, 4).End(xlToRight))
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_var.Range(ws_var.Cells(26, 4), ws_var.Cells(27, 4))
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0%"
End With

With ws_var.Range(ws_var.Cells(26, 5), ws_var.Cells(27, 9))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0.0000%"
End With

With ws_var.Range(ws_var.Cells(33, 3), ws_var.Cells(33, 3).End(xlToRight))
    .Interior.Color = RGB(0, 32, 96)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

With ws_var.Range(ws_var.Cells(34, 3), ws_var.Cells(34, 3).End(xlDown))
    .Font.Bold = True
    .HorizontalAlignment = xlHAlignLeft
    .VerticalAlignment = xlVAlignCenter

End With

With ws_var.Range(ws_var.Cells(34, 4), ws_var.Cells(34, 4).End(xlDown))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0.000%"
End With

With ws_var.Range(ws_var.Cells(34, 5), ws_var.Cells(34, 5).End(xlDown))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .NumberFormat = "0.0%"
End With

With ws_var.Range(ws_var.Cells(34, 6), ws_var.Cells(34, 6).End(xlDown))
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
    .Font.Bold = True
End With

With ws_var.Range(ws_var.Cells(34, 7), ws_var.Cells(34, 7).End(xlDown))
    .FormatConditions.Delete
End With

With ws_var.Range(ws_var.Cells(34, 7), ws_var.Cells(34, 7).End(xlDown)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Passed")
    .Font.Color = RGB(0, 176, 80)
    .Font.Bold = True
End With

With ws_var.Range(ws_var.Cells(34, 7), ws_var.Cells(34, 7).End(xlDown)).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="Failed")
    .Font.Color = RGB(192, 0, 0)
    .Font.Bold = True
End With

End Sub