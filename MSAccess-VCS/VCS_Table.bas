Option Compare Database

Option Private Module
Option Explicit


Public Sub ReloadTable(ByVal tbl_name As String)
    ImportTableDef tbl_name & ".xml"
    VCS_ExportTableDef tbl_name, VCS_SourcePath & "tbldef\"
End Sub


Public Sub VCS_ExportLinkedTable(ByVal tbl_name As String, ByVal obj_path As String)
    On Error GoTo Err_LinkedTable
    
    Dim tempFilePath As String
    Dim db As DAO.Database
    Set db = CurrentDb
    
    tempFilePath = VCS_File.VCS_TempFile()
    
    Dim FSO As Object
    Dim OutFile As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    VCS_Dir.VCS_MkDirIfNotExist obj_path
    
    Set OutFile = FSO.CreateTextFile(tempFilePath, overwrite:=True, Unicode:=True)
    
    OutFile.Write db.TableDefs(tbl_name).name
    OutFile.Write vbCrLf
    
    If InStr(1, db.TableDefs(tbl_name).connect, "DATABASE=" & CurrentProject.path) Then
        'change to relatave path
        Dim connect() As String
        connect = Split(db.TableDefs(tbl_name).connect, CurrentProject.path)
        OutFile.Write connect(0) & "." & connect(1)
    Else
        OutFile.Write db.TableDefs(tbl_name).connect
    End If
    
    OutFile.Write vbCrLf
    OutFile.Write db.TableDefs(tbl_name).SourceTableName
    OutFile.Write vbCrLf
    
    Dim td As DAO.TableDef
    Set td = db.TableDefs(tbl_name)
    Dim idx As DAO.index
    
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

Private Sub InsertTableField(ByRef TheTable As TableDef, ByRef TheField As Field, ByVal Position As Integer)
    Dim oField As Field
    
    For Each oField In TheTable.Fields
        If oField.OrdinalPosition >= Position Then
            oField.OrdinalPosition = oField.OrdinalPosition + 1
        End If
    Next
    
    TheTable.Fields.Refresh
    
    TheField.OrdinalPosition = Position
    TheTable.Fields.Append TheField
    TheTable.Fields.Refresh
    
End Sub

' Save a Table Definition as SQL statement
Public Sub VCS_ExportTableDef(ByVal TableName As String, ByVal directory As String)
    Dim fileName As String
    Dim db As Database
    fileName = directory & TableName & ".xml"
    
    ' If JSON version exists then use template model as well.
    If VCS_FileExists(directory & TableName & ".json") Then
        Dim FSO As New FileSystemObject
        Dim JsonTS As TextStream
        Dim JsonText As String
        Dim Parsed As Dictionary
        Dim oDict As New Dictionary
        Dim oFieldsDict As New Dictionary
        Dim oIndexesDict As New Dictionary
        Dim oIndex As index
        Dim oField As Field
        Dim oOtherIndex As index
        Dim oOtherField As Field
        Dim oTblDef As Object
        Dim oTemplateTblDef As TableDef
        Dim iFieldIndex As Integer
        
        Set db = CurrentDb
        ' Read .json file
        Set JsonTS = FSO.OpenTextFile(directory & TableName & ".json", ForReading)
        JsonText = JsonTS.ReadAll
        JsonTS.Close
        
        ' Parse json to Dictionary
        ' "values" is parsed as Collection
        ' each item in "values" is parsed as Dictionary
        Set Parsed = VCS_JsonConverter.ParseJson(JsonText)
        
        oDict("Template") = Parsed("Template")
        
        Set oTblDef = db.TableDefs(TableName)
        Set oTemplateTblDef = db.TableDefs(Parsed("Template"))
        
        ' Compare Fields with template table.
        ' If there are new Fields then they should be moved to the
        ' template table as it should be the superset of valid data.
        iFieldIndex = 0
        For Each oField In oTblDef.Fields
            Set oOtherField = Nothing
            
            On Error Resume Next
            Set oOtherField = oTemplateTblDef.Fields(oField.Name)
            
            If oOtherField Is Nothing Then
            
                If VCS_UpdateTemplateTable Then
                    iFieldIndex = iFieldIndex + 1
                    Debug.Print "Adding missing Field[" & oField.Name & "] @" & iFieldIndex & " from Table[" & oTemplateTblDef.Name & "]"
                    Set oOtherField = oTemplateTblDef.CreateField(oField.Name, oField.Type, oField.Size)
                    InsertTableField oTemplateTblDef, oOtherField, iFieldIndex
                Else
                    ' Not updating template, but extra fields exist.
                    Debug.Print "WARNING:  Template Table[" & oTemplateTblDef.Name & "] is missing Field[" & oField.Name & "]"
                End If
                
            Else
                iFieldIndex = oOtherField.OrdinalPosition
            End If
            
        Next oField
        On Error GoTo 0
        
        ' Compare indexes with template table.
        ' If there are new indexes then they should be moved to the
        ' template table as it should be the superset of valid data.
        For Each oIndex In oTblDef.Indexes
            Set oOtherIndex = Nothing
            
            On Error Resume Next
            Set oOtherIndex = oTemplateTblDef.Indexes(oIndex.Name)
            
            If oOtherIndex Is Nothing Then
                If VCS_UpdateTemplateTable Then
                    ' TODO:  Implement index copy.
                Else
                    Debug.Print "WARNING:  Template Table[" & oTemplateTblDef.Name & "] is missing index[" & oIndex.Name & "]"
                End If
            End If
            
        Next oIndex
        On Error GoTo 0
        
        ' Create dropped Indexes list.
        For Each oIndex In oTemplateTblDef.Indexes
            Set oOtherIndex = Nothing
            
            On Error Resume Next
            Set oOtherIndex = oTblDef.Indexes(oIndex.Name)
            
            If oOtherIndex Is Nothing Then
                oIndexesDict(oIndex.Name) = True
            End If
            
        Next oIndex
        On Error GoTo 0
        
        Set oDict("DropIndexes") = oIndexesDict.Keys
                
        ' Create dropped Fields list.
        For Each oField In oTemplateTblDef.Fields
            Set oOtherField = Nothing
            
            On Error Resume Next
            Set oOtherField = oTblDef.Fields(oField.Name)
            
            If oOtherField Is Nothing Then
                oFieldsDict(oField.Name) = True
            End If
            
        Next oField
        On Error GoTo 0
        
        Set oDict("DropFields") = oFieldsDict.Keys
        
        ' Write the template file.
        Dim oFSO As Object
        Dim oFile As Object
        
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        Set oFile = FSO.CreateTextFile(directory & TableName & ".json")
        oFile.WriteLine ConvertToJson(oDict, 4)
        oFile.Close
        Set oFSO = Nothing
        Set oFile = Nothing
        
    End If

    Application.ExportXML _
               ObjectType:=acExportTable, _
               DataSource:=TableName, _
               SchemaTarget:=fileName, _
               OtherFlags:=acExportAllTableAndFieldProperties

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
    Dim db As DAO.Database
    Dim Found As Boolean
    Dim Test As String
    
    Const NAME_NOT_IN_COLLECTION As Integer = 3265
    
    ' Assume the table or query does not exist.
    Found = False
    Set db = CurrentDb()
    
    ' Trap for any errors.
    On Error Resume Next
     
    ' See if the name is in the Tables collection.
    Test = db.TableDefs(TName).Name
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
        VCS_String.VCS_Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next
    
    VCS_String.VCS_Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    
    Count = 0
    For Each fieldObj In rs.Fields
        DoEvents
        If Count > 0 Then VCS_String.VCS_Sb_Append sb, ", "
        VCS_String.VCS_Sb_Append sb, "[" & fieldObj.Name & "]"
        Count = Count + 1
    Next

    TableExportSql = VCS_String.VCS_Sb_Get(sb)
End Function

Private Function CreateDOM()
    Dim dom
    Set dom = New DOMDocument60
    dom.async = False
    dom.validateOnParse = False
    dom.resolveExternals = False
    Set CreateDOM = dom
End Function

' Export the lookup table `tblName` to `source\tables`.
Public Sub VCS_ExportTableData(ByVal tbl_name As String, ByVal obj_path As String)
    Dim FSO As Object
    Dim OutFile As Object
    Dim rs As DAO.Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim c As Long, Value As Variant
    Dim oXMLFile As Object
    Dim oXSLFile As Object
    Dim xmlElement As Object
    Dim fileName As String
    Dim FieldName As String
    
    Dim doc, xsl, out As DOMDocument60
    Dim xslt As XSLTemplate60
    Dim str As String
    Dim xslproc
    
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    fileName = obj_path & tbl_name & ".xml"
    
    Application.ExportXML ObjectType:=acExportTable, DataSource:=tbl_name, DataTarget:=fileName, OtherFlags:=acEmbedSchema

    Set doc = CreateDOM
    doc.Load (fileName)
    
    Set xmlElement = doc.SelectSingleNode("/root/dataroot")
    If Not xmlElement Is Nothing Then
        xmlElement.removeAttribute ("generated")
    End If
    
    doc.Save fileName
    
End Sub

Public Sub VCS_ImportLinkedTable(ByVal tblName As String, ByRef obj_path As String)
    Dim db As DAO.Database
    Dim FSO As Object
    Dim InFile As Object
    
    Set db = CurrentDb
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
    Set td = db.CreateTableDef(InFile.ReadLine())
    
    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then 'replace relative path with literal path
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & CurrentProject.path & "\")
    End If
    td.Attributes = dbAttachSavePWD
    td.connect = connect
    
    td.SourceTableName = InFile.ReadLine()
    db.TableDefs.Append td
    
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
    Dim vField As Variant
    Dim SQL As String
    SQL = "CREATE INDEX __uniqueindex ON " & td.Name & " ("
    
    For Each vField In Split(Fields, ";+")
        SQL = SQL & "[" & vField & "]" & ","
    Next
    'remove extraneous comma
    SQL = Left$(SQL, Len(SQL) - 1)
    
    SQL = SQL & ") WITH PRIMARY"
    CurrentDb.Execute SQL
    
Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
    
End Sub

' Import Table Definition
Public Sub VCS_ImportTableDef(ByVal tblName As String, ByVal directory As String, Optional ByVal exclude_templates As Boolean = False)
    Dim filePath As String
    Dim tbl As Object
    Dim prefix As String
    Dim sTmpTable As String
    Dim oFields As Object
    Dim vVal As Variant
    Dim db As Database
    
    Set db = CurrentDb
    
    If exclude_templates And VCS_FileExists(directory & tblName & ".json") Then
        Exit Sub
    End If

    ' Drop table first.
    On Error GoTo Err_MissingTable
    Set tbl = db.TableDefs(tblName)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        If tblName <> "USysRibbonImages" And DCount("*", "[" & tblName & "]") > 0 Then
            sTmpTable = "tmp_" & tblName
            DoCmd.CopyObject , sTmpTable, acTable, tblName
        End If
        db.Execute "Drop Table [" & tblName & "]"
    End If
    
    ' If JSON version exists then use template model instead.
    If VCS_FileExists(directory & tblName & ".json") Then
        Dim FSO As New FileSystemObject
        Dim JsonTS As TextStream
        Dim JsonText As String
        Dim Parsed As Dictionary
        
        ' Read .json file
        Set JsonTS = FSO.OpenTextFile(directory & tblName & ".json", ForReading)
        JsonText = JsonTS.ReadAll
        JsonTS.Close
        
        ' Parse json to Dictionary
        ' "values" is parsed as Collection
        ' each item in "values" is parsed as Dictionary
        Set Parsed = VCS_JsonConverter.ParseJson(JsonText)
        
        If Not TableExists(Parsed("Template")) Then
            VCS_ImportTableDef Parsed("Template"), directory
        End If
        
        DoCmd.CopyObject , tblName, acTable, Parsed("Template")
        db.TableDefs.Refresh
        If Parsed.Exists("DropIndexes") Then
            For Each vVal In Parsed("DropIndexes")
                db.TableDefs(tblName).Indexes.Delete vVal
            Next vVal
        End If
       
        If Parsed.Exists("DropFields") Then
            For Each vVal In Parsed("DropFields")
                db.TableDefs(tblName).Fields.Delete vVal
            Next vVal
        End If
        
    Else
        filePath = directory & tblName & ".xml"
        Application.ImportXML DataSource:=filePath, ImportOptions:=acStructureOnly
    End If
    
    prefix = Left(tblName, 2)
    If prefix = "t_" Or prefix = "u_" Then
        Application.SetHiddenAttribute acTable, tblName, True
    End If
    
    If sTmpTable <> "" Then
        db.Execute "INSERT INTO [" & tblName & "] SELECT * FROM [" & sTmpTable & "]"
        db.Execute "Drop Table [" & sTmpTable & "]"
    End If
    
    Exit Sub
    
Err_MissingTable:
    ' Nothing to do here
    Resume Next
End Sub

' Import the lookup table `tblName` from `source\tables`.
Public Sub VCS_ImportTableData(ByVal tblName As String, ByVal obj_path As String, Optional ByVal appendOnly As Boolean = False)
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim FSO As Object
    Dim InFile As Object
    Dim c As Long, buf As String, Values() As String, Value As Variant
    
    Set db = CurrentDb

    If Not (appendOnly) Then
        ' Don't delete existing data
        db.Execute "DELETE FROM [" & tblName & "];"
    End If
    
    Application.ImportXML DataSource:=obj_path, ImportOptions:=acAppendData

End Sub


