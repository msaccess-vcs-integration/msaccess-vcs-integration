Attribute VB_Name = "VCS_Table"
Option Compare Database

Option Private Module
Option Explicit





Public Sub VCS_ExportLinkedTable(ByVal tbl_name As String, ByVal obj_path As String)
    On Error GoTo Err_LinkedTable
    
    Dim tempFilePath As String
    
    tempFilePath = VCS_File.VCS_TempFile()
    
    Dim FSO As Object
    Dim OutFile As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.VCS_MkDirIfNotExist obj_path
    
    Set OutFile = FSO.CreateTextFile(tempFilePath, overwrite:=True, Unicode:=True)
    
    OutFile.Write CurrentDb.TableDefs(tbl_name).name
    OutFile.Write vbCrLf
    
    If InStr(1, CurrentDb.TableDefs(tbl_name).connect, "DATABASE=" & CurrentProject.Path) Then
        'change to relatave path
        Dim connect() As String
        connect = Split(CurrentDb.TableDefs(tbl_name).connect, CurrentProject.Path)
        OutFile.Write connect(0) & "." & connect(1)
    Else
        OutFile.Write CurrentDb.TableDefs(tbl_name).connect
    End If
    
    OutFile.Write vbCrLf
    OutFile.Write CurrentDb.TableDefs(tbl_name).SourceTableName
    OutFile.Write vbCrLf
    
    Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim td As DAO.TableDef
    Set td = Db.TableDefs(tbl_name)
    Dim idx As DAO.Index
    
    For Each idx In td.Indexes
        If idx.Primary Then
            OutFile.Write Right$(idx.Fields, Len(idx.Fields) - 1)
            OutFile.Write vbCrLf
        End If

    Next
    
Err_LinkedTable_Fin:
    On Error Resume Next
    OutFile.Close
    'save files as .odbc
    VCS_File.VCS_ConvertUcs2Utf8 tempFilePath, obj_path & tbl_name & ".LNKD"
    
    Exit Sub
    
Err_LinkedTable:
    OutFile.Close
    MsgBox Err.Description, vbCritical, "ERROR: EXPORT LINKED TABLE"
    Resume Err_LinkedTable_Fin
End Sub

' Save a Table Definition as SQL statement
Public Sub VCS_ExportTableDef(ByVal TableName As String, ByVal directory As String)
    Dim fileName As String
    fileName = directory & TableName & ".xml"
    
    Application.ExportXML _
               ObjectType:=acExportTable, _
               DataSource:=TableName, _
               SchemaTarget:=fileName

    'export Data Macros
    VCS_DataMacro.VCS_ExportDataMacros TableName, directory
End Sub


' Determine if a table or exists.
' based on sample code of support.microsoftcom
' ARGUMENTS:
'    TName: The name of a table or query.
'
' RETURNS: True (it exists) or False (it does not exist).
Private Function TableExists(ByVal TName As String) As Boolean
    Dim Db As DAO.Database
    Dim Found As Boolean
    Dim Test As String
    
    Const NAME_NOT_IN_COLLECTION As Integer = 3265
    
     ' Assume the table or query does not exist.
    Found = False
    Set Db = CurrentDb()
    
     ' Trap for any errors.
    On Error Resume Next
     
     ' See if the name is in the Tables collection.
    Test = Db.TableDefs(TName).name
    If Err.Number <> NAME_NOT_IN_COLLECTION Then Found = True
    
    ' Reset the error variable.
    Err = 0
    
    TableExists = Found
End Function

' Build SQL to export `tbl_name` sorted by each field from first to last
Private Function TableExportSql(ByVal tbl_name As String) As String
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, Count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)
    
    sb = VCS_String.VCS_Sb_Init()
    VCS_String.VCS_Sb_Append sb, "SELECT "
    
    Count = 0
    For Each fieldObj In rs.Fields
        If Count > 0 Then VCS_String.VCS_Sb_Append sb, ", "
        VCS_String.VCS_Sb_Append sb, "[" & fieldObj.name & "]"
        Count = Count + 1
    Next
    
    VCS_String.VCS_Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    
    Count = 0
    For Each fieldObj In rs.Fields
        DoEvents
        If Count > 0 Then VCS_String.VCS_Sb_Append sb, ", "
        VCS_String.VCS_Sb_Append sb, "[" & fieldObj.name & "]"
        Count = Count + 1
    Next

    TableExportSql = VCS_String.VCS_Sb_Get(sb)
End Function

' Export the lookup table `tblName` to `source\tables`.
Public Sub VCS_ExportTableData(ByVal tbl_name As String, ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim rs As DAO.Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim c As Long, Value As Variant
    Dim oXMLFile As Object
    Dim xmlElement As Object
    Dim fileName As String
    
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    fileName = obj_path & tbl_name & ".xml"
    
    Application.ExportXML ObjectType:=acExportTable, DataSource:=tbl_name, DataTarget:=fileName, OtherFlags:=acEmbedSchema
    
    ' Remove the generated date field to make diff easier.
    Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    oXMLFile.async = False
    oXMLFile.validateOnParse = False
    oXMLFile.Load (fileName)
    
    Set xmlElement = oXMLFile.SelectSingleNode("/root/dataroot")
    If Not xmlElement Is Nothing Then
        xmlElement.removeAttribute ("generated")
        oXMLFile.Save (fileName)
    End If

End Sub

Public Sub VCS_ImportLinkedTable(ByVal tblName As String, ByRef obj_path As String)
    Dim Db As DAO.Database
    Dim FSO As Object
    Dim InFile As Object
    
    Set Db = CurrentDb
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFilePath As String
    tempFilePath = VCS_File.VCS_TempFile()
    
    VCS_ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFilePath, iomode:=ForReading, create:=False, Format:=TristateTrue)
    
    On Error GoTo err_notable:
    DoCmd.DeleteObject acTable, tblName
    
    GoTo err_notable_fin
    
err_notable:
    Err.Clear
    Resume err_notable_fin
    
err_notable_fin:
    On Error GoTo Err_CreateLinkedTable:
    
    Dim td As DAO.TableDef
    Set td = Db.CreateTableDef(InFile.ReadLine())
    
    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then 'replace relative path with literal path
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & CurrentProject.Path & "\")
    End If
    td.Attributes = dbAttachSavePWD
    td.connect = connect
    
    td.SourceTableName = InFile.ReadLine()
    Db.TableDefs.Append td
    
    GoTo Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin
    
Err_CreateLinkedTable_Fin:
    'this will throw errors if a primary key already exists or the table is linked to an access database table
    'will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:
    
    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim Field As Variant
    Dim sql As String
    sql = "CREATE INDEX __uniqueindex ON " & td.name & " ("
    
    For Each Field In Split(Fields, ";+")
        sql = sql & "[" & Field & "]" & ","
    Next
    'remove extraneous comma
    sql = Left$(sql, Len(sql) - 1)
    
    sql = sql & ") WITH PRIMARY"
    CurrentDb.Execute sql
    
Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
    
End Sub

' Import Table Definition
Public Sub VCS_ImportTableDef(ByVal tblName As String, ByVal directory As String)
    Dim filePath As String
    Dim tbl As Object
    
    ' Drop table first.
    On Error Resume Next
    Set tbl = CurrentDb.TableDefs(tblName)
    If Not tbl Is Nothing Then
        CurrentDb.Execute "Drop Table [" & tblName & "]"
    End If
    filePath = directory & tblName & ".xml"
    Application.ImportXML DataSource:=filePath, ImportOptions:=acStructureOnly

End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub VCS_ImportTableData(ByVal tblName As String, ByVal obj_path As String)
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO As Object
    Dim InFile As Object
    Dim c As Long, buf As String, Values() As String, Value As Variant
    
    Set Db = CurrentDb
    
    Db.Execute "DELETE FROM [" & tblName & "];"
    
    Application.ImportXML DataSource:=obj_path & tblName & ".xml", ImportOptions:=acAppendData

End Sub
