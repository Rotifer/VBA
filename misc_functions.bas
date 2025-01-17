Option Explicit

' Run the CLEAN function over a sheet used range and overwrite the original with
'  the CLEANed version.
Sub CleanCells()
    Dim sh As Worksheet
    Dim usedRng As Range
    Dim cell As Range
    Set sh = ActiveWorkbook.Worksheets("Sheet1")
    Set usedRng = sh.UsedRange
    MsgBox usedRng.Address
    For Each cell In usedRng.Cells
        cell.Value = Application.WorksheetFunction.Clean(cell.Value)
    Next cell
End Sub
