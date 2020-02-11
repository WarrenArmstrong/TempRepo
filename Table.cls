VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Internal collection to store dictionary objects
Private objColl As Collection

'executes when an object of the class is initialized
'instantiates the collection of dictionaries
Private Sub Class_Initialize()
    Set objColl = New Collection
End Sub

'Get method for the encapsulated collection
Public Property Get internal() As Collection
    Set internal = objColl
End Property

'Returns a deep copy of the object, can pass a boolean
'value to only return an empty table
Public Function clone(Optional makeEmpty As Boolean) As Table
    Dim newRow As Object
    Dim key As Variant
    Dim i As Long
    
    Set clone = New Table
    If Not makeEmpty Then
        For i = 1 To objColl.Count
            Set newRow = clone.newRow(True)
            For Each key In objColl(i).keys
                newRow.Add key, objColl(i)(key)
            Next key
        Next i
    End If
End Function

'Filter method that returns a new table object populated
'with all dictionaries that have the parameter expression
'evaluate to true
Public Function where(lambda As String) As Table
    Set where = Me.clone(False)
    
    Dim i As Long
    For i = where.internal.Count To 1 Step -1
        If Not evaluateLambda(lambda, i) Then
            where.internal.Remove (i)
        End If
    Next i
End Function

'Filter method that returns a table with the first dictionary
'to evaluate to true with the parameter expression. If
'no dictionaries have this property an empty table
'is returned
Public Function singleOrDefault(lambda As String) As Table
    Dim newRow As Object
    Dim key As Variant
    
    Set singleOrDefault = Me.clone(True)
    
    Dim i As Long
    For i = 1 To objColl.Count
        If evaluateLambda(lambda, i) Then
            Set newRow = singleOrDefault.newRow(True)
            For Each key In objColl(i).keys
                newRow.Add key, objColl(i)(key)
            Next key
            Exit Function
        End If
    Next i
End Function

'Returns a dictionary object with the keys associated
'with this table and nulls as the associated values
'can return an empty dictionary object if the parameter is true
Public Function newRow(Optional makeEmpty As Boolean) As Object
    Set newRow = Me.dictObj
    
    If objColl.Count <> 0 And Not makeEmpty Then
        Dim key As Variant
        For Each key In objColl(1).keys
            newRow.Add key, Null
        Next key
    End If
    
    objColl.Add newRow
End Function








'helper method for returning the evaluation of an
'expression for a certian dictionary in the table
Private Function evaluateLambda(ByVal lambda As String, ByVal row As Long) As Variant
    evaluateLambda = Evaluate(replaceLiterals("=" & lambda, row))
End Function

'helper function for the above function
Private Function replaceLiterals(ByVal lambda As String, ByVal row As Long) As String
    replaceLiterals = lambda
    
    Dim key As Variant
    For Each key In objColl(row)
        replaceLiterals = Replace(lambda, "{" & key & "}", objColl(row)(key))
    Next key
End Function

'method that takes the name of a table in the workbook
'and returns a table object representing that wb table
Public Function readTable(TableName As String) As Table
    Set readTable = readRange(Range(TableName & "[#All]"))
End Function

'method that takes a range object and returns a table
'object representing that range
Public Function readRange(rng As Range) As Table
    Set readRange = readArray(rng.Value2)
End Function


'method that converts a 2d array to a collection
'of dictionaries. Returns a table object representing this structure
Public Function readArray(arr As Variant) As Table
    Set objColl = New Collection
    
    Dim i As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        objColl.Add dictFromRow(arr, i)
    Next i
    
    Set readArray = Me
End Function


'helper method that returns a dictionary representing a certian row
'of a 2d array
Private Function dictFromRow(ByRef arr As Variant, ByVal row As Long) As Object
    Set dictFromRow = dictObj
    
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        dictFromRow.Add arr(LBound(arr, 1), i), arr(row, i)
    Next i
End Function

'method that converts the current table object into a 2d array
Public Function toArray() As Variant
    Dim output As Variant
    ReDim output(1 To objColl.Count + 1, 1 To objColl(1).Count)
    
    Dim i As Long, j As Long
    
    Dim key As Variant
    For Each key In objColl(1).keys
        j = j + 1
        output(1, j) = key
    Next key
    
    For i = LBound(output, 1) + 1 To UBound(output, 1)
        For j = 1 To UBound(output, 2)
            output(i, j) = objColl(i - 1)(output(1, j))
        Next j
    Next i
    
    toArray = output
End Function

'method that writes the current table object
'to a range in the workbook with the topleft of
'the range specified by the parameter range object
Public Function toRange(topLeft As Range) As Range
    Dim output As Variant
    output = toArray()
    
    topLeft.Resize(UBound(output, 1), UBound(output, 2)).Value2 = output
    Set toRange = topLeft.Resize(UBound(output, 1), UBound(output, 2))
End Function

'method that replaces the workbook table with the name
'matching the parameter string with a representation
'of the current table object
Public Function toTable(TableName As String)
    Dim oldRange As Range: Set oldRange = Range(TableName & "[#All]")
    oldRange.ClearContents
    oldRange.Worksheet.ListObjects.Add(xlSrcRange, toRange(oldRange), , xlYes).Name = TableName
End Function
















Public Function dictObj() As Object
    Set dictObj = CreateObject("Scripting.Dictionary")
End Function


