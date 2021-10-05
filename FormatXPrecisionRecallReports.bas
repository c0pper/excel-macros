Attribute VB_Name = "Module1"
Sub FormatXPrecisionRecallReports()
Attribute FormatXPrecisionRecallReports.VB_Description = "Format Precision and Recall reports from studio"
Attribute FormatXPrecisionRecallReports.VB_ProcData.VB_Invoke_Func = "R\n14"
'
' FormatPrecisionRecallReports Macro
' Format Extraction Precision and Recall reports from studio
'
' Keyboard Shortcut: Ctrl+Shift+R
'
    Columns("D:M").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:I").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").ColumnWidth = 84.43
    Columns("A:A").ColumnWidth = 93
    Columns("F:F").ColumnWidth = 14.71
    Columns("D:F").Select
    Selection.NumberFormat = "0.00"
    Columns("A:F").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A:$F"), , xlYes).Name = _
        "Table2"
    Columns("A:F").Select
    ActiveSheet.ListObjects("Table2").TableStyle = "TableStyleLight9"
    Columns("C:C").Select
    Selection.ListObject.ListColumns(3).Delete
    Range("Table2[[X-Precision]:[X-F-Measure]]").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0.7"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.4", Formula2:="=0.7"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.1", Formula2:="=0.4"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
