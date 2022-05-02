Sub format_cells()
'
' format_cells Macro
'
' Keyboard Shortcut: Ctrl+Shift+F

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


