Attribute VB_Name = "Module2"
Sub format_cells()
Attribute format_cells.VB_Description = "jsdkfhjul"
Attribute format_cells.VB_ProcData.VB_Invoke_Func = "a\n14"
'
' format_cells Macro
' jsdkfhjul
'
' Keyboard Shortcut: Ctrl+a
'
    Columns("A:A").Select
    Selection.ColumnWidth = 10
    Columns("B:D").Select
    Selection.ColumnWidth = 40
    Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Range("B2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Sub
