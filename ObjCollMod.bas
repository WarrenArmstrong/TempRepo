Attribute VB_Name = "ObjCollMod"
Option Explicit

Public Function readTable(TableName As String) As Collection
    Set readTable = readRange(Range(TableName & "[#All]"))
End Function

Public Function readRange(rng As Range) As Collection
    Set readRange = readArray(rng.Value2)
End Function

Public Function readArray(arr As Variant) As Collection
    Set readArray = New Collection
    
    Dim i As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        readArray.Add dictFromRow(arr, i)
    Next i
End Function

Private Function dictFromRow(ByRef arr As Variant, ByVal row As Long) As Object
    Set dictFromRow = dictObj
    
    Dim i As Long
    For i = LBound(arr, 2) To UBound(arr, 2)
        dictFromRow.Add arr(LBound(arr, 1), i), arr(row, i)
    Next i
End Function

Public Function dictObj() As Object
    Set dictObj = CreateObject("Scripting.Dictionary")
End Function


Public Function toArray(objColl As Collection) As Variant
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

Public Function toRange(objColl As Collection, rng As Range)
    Dim output As Variant
    output = toArray(objColl)
    
    rng.Resize(UBound(output, 1), UBound(output, 2)).Value2 = output
End Function

Public Function toTable(objColl As Collection, TableName As String)
    toRange objColl, Range(TableName & "[#All]")
End Function

