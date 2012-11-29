VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsKeyValueMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' SUMMARY:
'A simple example of object-oriented programming in VBA
'Provides a class to implement a simple dictionary where uninque string keys map to string values.
'Recent Excel versions on Windows provide a Dictionary object.
'This class is NOT TO BE USED when running recent versions of Excel on Windows. Use the provided Dictionary object instead!
'See Microsoft article: http://support.microsoft.com/kb/187234
'CODE NOTES:
'Internally the class uses a Collection and stores the key-value pairs as two-element string arrays.
'Its underlying Collection basis hidden from calling code and its public methods and properties
'try to make it look like a dictionary/hash.
'Both keys and values are simple string scalars so keys cannot map to arrays.  Workarounds are possible where the values can be
'concatenated and then split into arrays when retrieved.
'If a key-pair is added to the collection for a key that already exists in the collection, the old entry is over-written.
'If different behaviour is required based on different values that have the same key, the calling code needs to deal with that.
'The keys are CASE-SENSITIVE ("a" != "A")
'EXAMPLE USAGE:
'    Option Explicit
'    Sub test_clsKeyValueMap()
'        Dim dict As clsKeyValueMap
'        Set dict = New clsKeyValueMap
'        Dim Keys() As String
'        Dim Values() As String
'        Dim i As Long
'
'        'Add values to the dict object
'        Call dict.addKeyValuePair("A", 1)
'        Call dict.addKeyValuePair("a", 2)
'        Call dict.addKeyValuePair("B", 3)
'        Call dict.addKeyValuePair("A", 4)
'
'        'Test methods and properties
'        MsgBox dict.KeyValueMapCount
'        MsgBox dict.HasKey("A")
'        MsgBox dict.HasKey("a")
'        MsgBox dict.HasKey("B")
'        MsgBox dict.ValueForKey("B")
'        Keys = dict.Keys
'        Values = dict.Values
'        MsgBox Join(Keys, ", ")
'        MsgBox Join(Values, ", ")
'
'
'        'Loop over the dict object printing keys and their values to the immediate window.
'        For i = 0 To UBound(Keys)
'            Debug.Print Keys(i), " => ", dict.ValueForKey(Keys(i))
'        Next i
'        'Output:
'        'a       =>    2
'        'B       =>    3
'        'A       =>    4
'    End Sub
'TO DO:
'Allow more complex data structures, dictionaries of arrays, dictionaries of dictionaries etc.
'TO USE:
'Import the class module to project
'Create a normal module and paste in the test code given above in the sub named "test_clsKeyValueMap()" to ensure it is running
'AUTHOR:
'mick@javascript-spreadsheet-programming.com

Private keyValueMap As Collection
Private Const CODEVERSION As String = "TEST"
'Constructor: Initialize an empty collection
Private Sub Class_Initialize()
    Set keyValueMap = New Collection
End Sub
'Return the number of elements in the collection instance
Public Property Get KeyValueMapCount()
    KeyValueMapCount = keyValueMap.Count
End Property
'A flag to providing the code version of the class
Public Property Get Version() As String
    Version = CODEVERSION
End Property
'Add a two element string array to the collection where the first element acts as the key and the second as the value
'If the key does not already exist in the collection add the pair,
'else the key already exists, remove the old key-value pair and add the new one.
Public Sub addKeyValuePair(key As String, value As String)
    Dim pair(1) As String
    
    pair(0) = key
    pair(1) = value
    If IndexForKey(key) < 0 Then
        keyValueMap.Add (pair)
    Else
        keyValueMap.Remove (IndexForKey(key))
        keyValueMap.Add (pair)
    End If
    
End Sub
'Given a key string, check the first elment of the array in each collection item for the key value.
'If a match is found, return the index where the match is found, else if not found return -1.
'This is a private property because the underlying collection is to be treated as unordered.
Private Property Get IndexForKey(key As String) As Long
    Dim keyValuePair() As String
    Dim i As Long
    
    IndexForKey = -1
    If KeyValueMapCount > 0 Then
        For i = 1 To KeyValueMapCount
            keyValuePair = keyValueMap(i)
            If keyValuePair(0) = key Then
                IndexForKey = i
                Exit Property
            End If
        Next i
    End If
End Property
'Return true if the given key is present in the collection
'Can be used by calling code to determine what to do if it needs to add a key-value pair where the key already exists.
Public Property Get HasKey(key As String) As Boolean
    HasKey = False
    If KeyValueMapCount > 0 And IndexForKey(key) > 0 Then
        HasKey = True
    End If
End Property
'Given a key check the collection for a match and return the matching value it maps to.
'If the key is not found, return a VB null string.
Public Property Get ValueForKey(key As String) As Variant
    Dim pair() As String
    
    If KeyValueMapCount > 0 And IndexForKey(key) > 0 Then
        pair = keyValueMap.item(IndexForKey(key))
        ValueForKey = pair(1)
    Else
        ValueForKey = vbNullString
    End If
    
End Property
'Return a string array of all the key values i.e. the first elements in the two-element array items.
'If the collection is empty, return an array of one element where that element is a VB null string.
'The keys array elements are in random order because they were retrieved using a For Each to loop over the collection
Public Property Get Keys() As String()
    Dim arrKeys() As String
    Dim i As Long: i = 0
    Dim item As Variant
    
    If KeyValueMapCount > 0 Then
        For Each item In keyValueMap
            ReDim Preserve arrKeys(i)
            arrKeys(i) = item(0)
            i = i + 1
        Next item
    Else
        ReDim Preserve arrKeys(i)
        arrKeys(i) = vbNullString
    End If
    
    Keys = arrKeys
    
End Property
'Return a string array of all the keyed values i.e. the second elements in the two-element array items.
'If the collection is empty, return an array of one element where that element is a VB null string.
'The values array elements are in random order because they were retrieved using a For Each to loop over the collection
Public Property Get Values() As String()
    Dim arrValues() As String
    Dim i As Long: i = 0
    Dim item As Variant
    
    If KeyValueMapCount > 0 Then
        For Each item In keyValueMap
            ReDim Preserve arrValues(i)
            arrValues(i) = item(1)
            i = i + 1
        Next item
    Else
        ReDim Preserve arrValues(i)
        arrValues(i) = vbNullString
    End If
    
    Values = arrValues
    
End Property
'Use the given a key to locate a matching item in the collection.
'If the key is found, remove that key-value pair from the collection and return "True" to indicate successful removal.
'If the key is not found, return "False".
Public Function DeleteKeyValuePair(key As String) As Boolean
    If KeyValueMapCount > 0 And IndexForKey(key) > 0 Then
        keyValueMap.Remove (IndexForKey(key))
        DeleteKeyValuePair = True
    Else
        DeleteKeyValuePair = False
    End If
End Function
