' Name all the cells of a sorted column using the unique values
'  of the starting cell column.
' Used to name blocks of ranges.
Sub AssignNamesToRange(rngForNames As Range, startCell As Range)
    Dim rngRowCount As Long: rngRowCount = rngForNames.Rows.Count
    Dim rngColCount As Long: rngColCount = rngForNames.Columns.Count
    Dim startCellColNum As Long: startCellColNum = startCell.Column
    Dim rngFirstColNum As Long: rngFirstColNum = rngForNames.Cells(1, 1).Column
    Dim rngLastColNum As Long: rngLastColNum = rngFirstColNum + rngColCount
    Dim leftColOffset As Long: leftColOffset = rngFirstColNum - startCellColNum
    Dim rightColOffset As Long: rightColOffset = rngLastColNum - startCellColNum
    Dim rngTopLeftCell As Range
    Dim rngBottomRightCell As Range
    Dim nextRowNum As Long
    Dim rngStart As Long
    Dim i As Long
    Dim j As Long
    Dim rngName As String
    Do While i <= rngRowCount
        j = 0
        Do While startCell.Offset(j, 0).Value = startCell.Value
            j = j + 1
        Loop
        rngName = startCell.Value
        Set rngTopLeftCell = startCell.Offset(0, leftColOffset)
        Set rngBottomRightCell = startCell.Offset(j - 1, rightColOffset - 1)
        Range(rngTopLeftCell, rngBottomRightCell).Name = rngName
        Set startCell = startCell.Offset(j, 0)
        If startCell.Row > rngRowCount Then Exit Do
        i = i + j
    Loop
End Sub

Sub RunAssignNamesToRange()
    Dim rngForNames As Range: Set rngForNames = ActiveWorkbook.Worksheets("Sheet1").UsedRange
    Dim startCell As Range: Set startCell = ActiveWorkbook.Worksheets("Sheet1").Range("C1")
    Call AssignNamesToRange(rngForNames, startCell)
End Sub
