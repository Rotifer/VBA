Attribute VB_Name = "SetUDFs"
Option Explicit
    ' Select Tools->References from the Visual Basic menu.
    ' Check box beside "Microsoft Scripting Runtime" in the list.


' Public function that can be called as a UDF.
' One glitch: The tab is disappearing (??)
' TODO: fix the tab
Public Function UNION(rng1 As Range, rng2 As Range) As Variant
    Dim set1 As clsSet: Set set1 = New clsSet
    Dim set2 As clsSet: Set set2 = New clsSet
    Dim unionDict As Dictionary
    Dim elements As Variant
    set1.InputRng = rng1
    set2.InputRng = rng2
    Set unionDict = set1.UNION(set2)
    UNION = Application.WorksheetFunction.Transpose(unionDict.Keys)
End Function

