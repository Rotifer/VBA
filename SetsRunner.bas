Attribute VB_Name = "SetsRunner"
Option Explicit
' Assumes the "Microsoft Scripting Runtime" reference has been added.
' HOWTO: Select Tools->References from the Visual Basic menu.
'  Check box beside "Microsoft Scripting Runtime" in the list.

' This module contains a set of subroutines the use the methos of clsSet
'  to perform set operations (UNION, INTERSECTION and DIFFERENCE) on values
'  extracted from pairs of ranges. The sets are implemented in the class as
'  dictionaries with the range values for each set as dictionary keys where the
'  values of each row are stored as tab-separated strings of the row cell values
'  for each range.

' I call the clsSet methods in UDFs that can be called directly in Excel formulas but
'  haven't been able to get that to work.

' ################################### subroutines #################################

' Given two ranges and a target cell for first output, write out
'  the UNION result.
Sub Union(rng1 As Range, rng2 As Range, startCell As Range)
    Dim set1 As clsSet: Set set1 = New clsSet
    Dim set2 As clsSet: Set set2 = New clsSet
    Dim unionDict As Dictionary
    
    set1.InputRng = rng1
    set2.InputRng = rng2
    Set unionDict = set1.Union(set2)
    Call PrintSet(ActiveCell, unionDict, vbTab)
End Sub
' Runner for UNION.
Sub CallUnion()
    Dim rng1 As Range, rng2 As Range
    Set rng1 = Worksheets("Set1").Range("A1:B10")
    Set rng2 = Worksheets("Set2").Range("A1:B10")
    Call Union(rng1, rng2, ActiveCell)
End Sub

' Print the split result of a concatenated keys of a dictionary (our Set),
'  beginning at the start cell using the split character argument to generate the output values
Sub PrintSet(startCell As Range, dictToPrint As Dictionary, splitChar As String)
    Dim i As Integer, j As Integer
    Dim key As Variant
    Dim arrRow() As String
    Dim cellVal As Va.riant
    i = 0
    j = 0
    For Each key In dictToPrint.Keys()
        arrRow = Split(key, splitChar)
        For j = 0 To UBound(arrRow)
            cellVal = arrRow(j)
            startCell.Offset(i, j).Value = cellVal
        Next j
        i = i + 1
    Next key
End Sub

' Given two ranges and a target cell for first output, write out
'  the INTERSECTION result.
Public Sub Intersection(rng1 As Range, rng2 As Range, startCell As Range)
    Dim set1 As clsSet: Set set1 = New clsSet
    Dim set2 As clsSet: Set set2 = New clsSet
    Dim intersectionDict As Dictionary
    
    set1.InputRng = rng1
    set2.InputRng = rng2
    Set intersectionDict = set1.Intersection(set2)
    Call PrintSet(ActiveCell, intersectionDict, vbTab)
End Sub

' Runner for INTERSECTION.
Sub CallIntersection()
    Dim rng1 As Range, rng2 As Range
    Set rng1 = Worksheets("Set1").Range("A1:B10")
    Set rng2 = Worksheets("Set2").Range("A1:B10")
    Call Intersection(rng1, rng2, ActiveCell)
End Sub

' Given two ranges and a target cell for first output, write out
'  the DIFFERENCE result.
Sub Difference(rng1 As Range, rng2 As Range, startCell As Range)
    Dim set1 As clsSet: Set set1 = New clsSet
    Dim set2 As clsSet: Set set2 = New clsSet
    Dim differenceDict As Dictionary
    
    set1.InputRng = rng1
    set2.InputRng = rng2
    Set differenceDict = set2.Difference(set1)
    Call PrintSet(ActiveCell, differenceDict, vbTab)
End Sub

' Runner for DIFFERENCE.
Sub CallDifference()
    Dim rng1 As Range, rng2 As Range
    Set rng1 = Worksheets("Set1").Range("A1:B10")
    Set rng2 = Worksheets("Set2").Range("A1:B10")
    Call Difference(rng1, rng2, ActiveCell)
End Sub
