VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_rng1 As Range
Private m_rng2 As Range
Private m_mySet1 As clsSet
Private m_mySet2 As clsSet
Private Const ws1Name As String = "Set1"
Private Const ws2Name As String = "Set2"

Private Sub Class_Initialize()
    Call DeleteSheets
    Call AddSheets
    Call AddValuesToSheets
    Set m_rng1 = ActiveWorkbook.Worksheets(ws1Name).Range("A1:B10")
    Set m_rng2 = ActiveWorkbook.Worksheets(ws2Name).Range("A1:B10")
    Set m_mySet1 = New clsSet
    Set m_mySet2 = New clsSet
    m_mySet1.InputRng = m_rng1
    m_mySet2.InputRng = m_rng2
End Sub
Private Sub DeleteSheets()
    On Error Resume Next
    Dim ws1 As Worksheet: Set ws1 = ActiveWorkbook.Worksheets(ws1Name)
    Dim ws2 As Worksheet: Set ws2 = ActiveWorkbook.Worksheets(ws2Name)
    Application.DisplayAlerts = False
    ws1.Delete
    ws2.Delete
    Application.DisplayAlerts = True
End Sub
Private Sub AddSheets()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = ActiveWorkbook.Worksheets().Add()
    Set ws2 = ActiveWorkbook.Worksheets().Add()
    ws1.Name = ws1Name
    ws2.Name = ws2Name
End Sub
Private Sub AddValuesToSheets()
    Dim arrMonths(12) As String
    Dim ws1 As Worksheet: Set ws1 = ActiveWorkbook.Worksheets(ws1Name)
    Dim ws2 As Worksheet: Set ws2 = ActiveWorkbook.Worksheets(ws2Name)
    Dim startRng1 As Range: Set startRng1 = ws1.Range("A1")
    Dim startRng2 As Range: Set startRng2 = ws2.Range("A1")
    Dim i As Integer
    arrMonths(0) = "Jan"
    arrMonths(1) = "Feb"
    arrMonths(2) = "Mar"
    arrMonths(3) = "Apr"
    arrMonths(4) = "May"
    arrMonths(5) = "Jun"
    arrMonths(6) = "Jul"
    arrMonths(7) = "Aug"
    arrMonths(8) = "Sep"
    arrMonths(9) = "Oct"
    arrMonths(10) = "Nov"
    arrMonths(11) = "Dec"
    For i = 0 To 9
        startRng1.Offset(i, 0).Value = arrMonths(i)
        startRng1.Offset(i, 1).Value = i + 1
    Next i
    For i = 2 To 11
        startRng2.Offset(i - 2, 0).Value = arrMonths(i)
        startRng2.Offset(i - 2, 1).Value = i + 1
    Next i
End Sub

Public Sub TestArrayValues()
    Dim rngValuesArray() As String: rngValuesArray = m_mySet1.rngValuesArray()
    Debug.Assert UBound(rngValuesArray) = 9
    Debug.Assert rngValuesArray(0) = "Jan" & vbTab & "1"
    Debug.Print "Method 'rngValuesArray' passes..."
End Sub

Public Sub TestRowsAsSet()
    Dim rowsAsSet As Dictionary: Set rowsAsSet = m_mySet1.rowsAsSet
    Debug.Assert rowsAsSet.Count = 10
    Debug.Print "Method 'rowsAsSet' passes...."
End Sub

Public Sub TestIsSuperset()
    Dim isSuperset_ As Boolean:
    isSuperset_ = m_mySet1.IsSuperset(m_mySet2)
    Debug.Assert isSuperset_ = False
    Debug.Print "Method 'IsSuperset' passes ...."
End Sub

Public Sub TestUnion()
    Dim union_ As Dictionary
    Set union_ = m_mySet1.UNION(m_mySet2)
    Debug.Assert union_.Count = 12
    Debug.Print "Method 'Union' passes ...."
End Sub

Public Sub TestIntersection()
    Dim intersection_ As Dictionary
    Set intersection_ = m_mySet1.Intersection(m_mySet2)
    Debug.Assert intersection_.Count = 8
    Debug.Assert Not intersection_.Exists("Jan" & vbTab & "1")
    Debug.Assert intersection_.Exists("Mar" & vbTab & "3")
    Debug.Print "Method 'Intersection' passes ...."
End Sub

Public Sub TestDifference()
    Dim difference_ As Dictionary
    Set difference_ = m_mySet1.Difference(m_mySet2)
    Debug.Assert difference_.Count = 2
    Debug.Assert difference_.Exists("Jan" & vbTab & "1") = True
    Debug.Assert difference_.Exists("Feb" & vbTab & "2") = True
    Debug.Print "Method 'Difference' passed ...."
End Sub
 
