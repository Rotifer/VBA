Attribute VB_Name = "modMain"
Option Explicit


Sub CopySheetUsedRng(pathToInputXLFile As String, _
              fromSheetName As String, _
              pathToOutputXLFile As String, _
              toSheetName As String)
    Dim inputWbk As Workbook
    Dim outputWbk As Workbook
    Dim fromSheet As Worksheet
    Dim toSheet As Worksheet
    Dim toCellRng As Range
    Dim toUsedRngRowCount As Long
    Set inputWbk = Application.Workbooks().Open(pathToInputXLFile)
    Set outputWbk = Application.Workbooks().Open(pathToOutputXLFile)
    Set fromSheet = inputWbk.Worksheets(fromSheetName)
    Set toSheet = outputWbk.Worksheets(toSheetName)
    toUsedRngRowCount = toSheet.UsedRange.Rows.Count
    If toSheet.Range("A1").Value = "" Then
        Set toCellRng = toSheet.Range("A1")
    Else
        Set toCellRng = toSheet.Range("A1").Offset(toUsedRngRowCount, 0)
    End If
    fromSheet.UsedRange.Copy toCellRng
    inputWbk.Close False
    outputWbk.Close True
End Sub


Sub RunCopySheetUsedRng()
    Dim pathToInputXLFile As String
    Dim pathToOutputXLFile As String: pathToOutputXLFile = "<path>\concatenated_file.xlsx"
    Dim fromSheetName As String: fromSheetName = "Sheet1"
    Dim toSheetName As String: toSheetName = "Sheet1"
    Dim i As Integer
    Application.ScreenUpdating = False
    For i = 1 To 609
        pathToInputXLFile = "<path_to_files>\table_" & i & ".xlsx"
        Call CopySheetUsedRng(pathToInputXLFile, fromSheetName, pathToOutputXLFile, toSheetName)
    Next i
    Application.ScreenUpdating = True
    MsgBox "Complete!"
End Sub
