Attribute VB_Name = "TestMod"
Option Explicit

Public Sub testTable()
    Dim people As Table: Set people = New Table
    
    people.readTable "People"
    
    Dim newRow As Object: Set newRow = people.newRow
    
    newRow("Name") = "Jeff"
    newRow("Age") = 23
    
    'Set people = people.clone(False)
    
    'Set people = people.where("{Age} = 21")
    
    'people.toRange Range("A1")
    
    
    Set people = people.where("{Age} = 21").toTable("Table2")
    people.toTable "Table2"
    'people.singleOrDefault("{Age} = 21").toTable "Table2"
End Sub
