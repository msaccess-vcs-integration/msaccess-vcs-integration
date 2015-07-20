Attribute VB_Name = "VCS_Relation"
Option Compare Database

Option Explicit

Public Sub ExportRelation(rel As DAO.Relation, filePath As String)
    Dim FSO, OutFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(filePath, True)

    OutFile.WriteLine rel.Attributes 'RelationAttributeEnum
    OutFile.WriteLine rel.name
    OutFile.WriteLine rel.table
    OutFile.WriteLine rel.foreignTable
    Dim f As DAO.Field
    For Each f In rel.Fields
        OutFile.WriteLine "Field = Begin"
        OutFile.WriteLine f.name
        OutFile.WriteLine f.ForeignName
        OutFile.WriteLine "End"
    Next
    OutFile.Close

End Sub


Public Sub ImportRelation(filePath As String)
    Dim FSO, InFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(filePath, 1)
    
    Dim rel As New DAO.Relation
    rel.Attributes = InFile.ReadLine
    rel.name = InFile.ReadLine
    rel.table = InFile.ReadLine
    rel.foreignTable = InFile.ReadLine
    Dim f As DAO.Field
    Do Until InFile.AtEndOfStream
        If "Field = Begin" = InFile.ReadLine Then
            Set f = New DAO.Field
            f.name = InFile.ReadLine
            f.ForeignName = InFile.ReadLine
            If "End" <> InFile.ReadLine Then
                Set f = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            rel.Fields.Append f
        End If
    Loop
    
    InFile.Close
    
    CurrentDb.Relations.Append rel

End Sub
