Attribute VB_Name = "modTest_clsSets"
Option Explicit
Sub testClass()
    Dim testCls As clsTestSet: Set testCls = New clsTestSet
    Call testCls.TestArrayValues
    Call testCls.TestRowsAsSet
    Call testCls.TestIntersection
    Call testCls.TestIsSuperset
    Call testCls.TestUnion
    ' Re-calling to test for mutation of set1, expect False
    Call testCls.TestIsSuperset
    Call testCls.TestDifference
End Sub

