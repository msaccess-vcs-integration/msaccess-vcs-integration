Option Compare Database

Option Explicit
' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
' Set to "*" to export the contents of all tables
'Only used in ExportAllSource
Private INCLUDE_TABLES As String
' This is used in ImportAllSource
Private DebugOutput As Boolean
'this is used in ExportAllSource
'Causes the VCS_ code to be exported
Private ArchiveMyself As Boolean

' Export configuration
Private ExportReports As Boolean
Private ExportQueries As Boolean
Private ExportForms As Boolean
Private ExportMacros As Boolean
Private ExportModules As Boolean
Private ExportTables As Boolean
'export/import all Queries as plain SQL text
Private HandleQueriesAsSQL As Boolean

Private Function StrToBool(ByVal sStr As String, ByVal default As Boolean)
    StrToBool = default

    sStr = LCase(sStr)
    If sStr = "yes" OR sStr = "true" Then
        StrToBool = True
    Else
        If sStr = "no" OR sStr = "false" Then
            StrToBool = False
        End If
    End If
End Function

Public Sub LoadCustomisations()
    Dim sValue As String
    Dim path As String

    path = VCS_Dir.VCS_ProjectPath()

    ' Load configuration items
    ArchiveMyself = StrToBool(GetSectionEntry("Config", "ArchiveMyself", path & "vcs.cfg"), False)
    INCLUDE_TABLES = GetSectionEntry("Config", "IncludeTables", path & "vcs.cfg")
    DebugOutput = StrToBool(GetSectionEntry("Config", "DebugOutput", path & "vcs.cfg"), False)
    ExportReports = StrToBool(GetSectionEntry("Config", "ExportReports", path & "vcs.cfg"), True)
    ExportForms = StrToBool(GetSectionEntry("Config", "ExportForms", path & "vcs.cfg"), True)
    ExportQueries = StrToBool(GetSectionEntry("Config", "ExportQueries", path & "vcs.cfg"), True)
    ExportMacros = StrToBool(GetSectionEntry("Config", "ExportMacros", path & "vcs.cfg"), True)
    ExportModules = StrToBool(GetSectionEntry("Config", "ExportModules", path & "vcs.cfg"), True)
    ExportTables = StrToBool(GetSectionEntry("Config", "ExportTables", path & "vcs.cfg"), True)
    HandleQueriesAsSQL = StrToBool(GetSectionEntry("Config", "HandleQueriesAsSQL", path & "vcs.cfg"), True)
    
End Sub

' NOTE:  VCS_ImportAllSource and VCS_ImportAllModules are in VCS_Loader
' This is because you can't replace modules while running code in those
' modules.
Public Sub VCS_ExportAllSource()
    LoadCustomisations
    ExportAllSource
End Sub
Public Sub VCS_ExportAllModules()
    LoadCustomisations
    ExportAllModules
End Sub
Public Sub VCS_ExportAllTableDefs()
    LoadCustomisations
    ExportAllTables doTableDefs:=True, doTableData:=False
End Sub
Public Sub VCS_ImportAllForms()
    LoadCustomisations
    CloseFormsReports
    ImportAllForms
End Sub
Public Sub VCS_ExportAllForms()
    DoCmd.Hourglass True
    LoadCustomisations
    ExportAllForms
    DoCmd.Hourglass False
End Sub
Public Sub VCS_ImportAllReports()
    LoadCustomisations
    CloseFormsReports
    ImportAllReports
End Sub
Public Sub VCS_ExportAllReports()
    LoadCustomisations
    ExportAllReports
End Sub
Public Sub VCS_ImportAllMacros()
    LoadCustomisations
    ImportAllMacros
End Sub
Public Sub VCS_ExportAllMacros()
    LoadCustomisations
    ExportAllMacros
End Sub
Public Sub VCS_ImportAllTableDefs()
    LoadCustomisations
    CloseFormsReports
    ImportAllTableDefs
End Sub
Public Sub VCS_ExportAllTableData()
    LoadCustomisations
    ExportAllTables doTableDefs:=False, doTableData:=True
End Sub
Public Sub VCS_ImportAllTableData()
    LoadCustomisations
    ImportAllTableData
    ImportAllTableDataMacros
End Sub


'returns true if named module is NOT part of the VCS code
Private Function IsNotVCS(ByVal moduleName As String) As Boolean
    If moduleName <> "VCS_ImportExport" And _
      moduleName <> "VCS_IE_Functions" And _
      moduleName <> "VCS_File" And _
      moduleName <> "VCS_Dir" And _
      moduleName <> "VCS_String" And _
      moduleName <> "VCS_Loader" And _
      moduleName <> "VCS_Table" And _
      moduleName <> "VCS_Reference" And _
      moduleName <> "VCS_DataMacro" And _
      moduleName <> "VCS_Report" And _
      moduleName <> "VCS_Relation" And _
      moduleName <> "VCS_Query" And _
      moduleName <> "VCS_IniHandler" And _
      moduleName <> "VCS_Button_Functions" Then
        IsNotVCS = True
    Else
        IsNotVCS = False
    End If

End Function

Public Sub ExportObjTypeSource(ByVal obj_type As Variant)
    Dim doc As Object ' DAO.Document
    Dim Db As Object
    Dim ucs2 As Boolean
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_path As String
    Dim obj_count As Integer

    Set Db = CurrentDb
    
    obj_type_split = Split(obj_type, "|")
    obj_type_label = obj_type_split(0)
    obj_type_name = obj_type_split(1)
    obj_type_num = Val(obj_type_split(2))
    obj_path = VCS_SourcePath & obj_type_label & "\"
    obj_count = 0
                
    If (obj_type_label = "forms" And ExportForms) _
       Or (obj_type_label = "reports" And ExportReports) _
       Or (obj_type_label = "macros" And ExportMacros) _
       Or (obj_type_label = "modules" And ExportModules) Then

        VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
        Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24)
        SysCmd acSysCmdInitMeter, "Exporting " & obj_type_label, Db.Containers(obj_type_name).Documents.Count
        For Each doc In Db.Containers(obj_type_name).Documents
            DoEvents
            If (Left$(doc.name, 1) <> "~") And _
               (IsNotVCS(doc.name) Or ArchiveMyself) Then
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = VCS_File.VCS_UsingUcs2
                End If
                VCS_IE_Functions.VCS_ExportObject obj_type_num, doc.name, obj_path & doc.name & ".bas", ucs2
                
                If obj_type_label = "reports" Then
                    VCS_Report.VCS_ExportPrintVars doc.name, obj_path & doc.name & ".pv"
                End If
                
                obj_count = obj_count + 1
                SysCmd acSysCmdUpdateMeter, obj_count
            End If
        Next
        SysCmd acSysCmdRemoveMeter

        Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15)
        SysCmd acSysCmdInitMeter, "Sanitizing " & obj_type_name, obj_count
        If obj_type_label <> "modules" Then
            VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
        End If
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If

    If (obj_type_label = "modules" And ExportModules) Then
        VCS_Reference.VCS_ExportReferences VCS_SourcePath
    End If
End Sub

Public Sub ExportAllModules()
    ExportObjTypeSource "modules|Modules|" & acModule
End Sub

Public Sub ExportAllForms()
    ExportObjTypeSource "forms|Forms|" & acForm
End Sub

Public Sub ExportAllReports()
    ExportObjTypeSource "reports|Reports|" & acReport
End Sub

Public Sub ExportAllMacros()
    ExportObjTypeSource "macros|Scripts|" & acMacro
End Sub

Public Sub ExportAllTables(Optional ByVal doTableDefs As Boolean = True, Optional ByVal doTableData As Boolean = True)
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean

    Set Db = CurrentDb
    source_path = VCS_SourcePath
    If ExportTables Then
        obj_path = source_path & "tables\"
        VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
        VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"
                
        Dim td As DAO.TableDef
        Dim tds As DAO.TableDefs
        Set tds = Db.TableDefs

        obj_type_label = "tbldef"
        obj_type_name = "Table_Def"
        obj_type_num = acTable
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        obj_data_count = 0
        VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))
                
        'move these into Table and DataMacro modules?
        ' - We don't want to determin file extensions here - or obj_path either!
        If doTableDefs Then
            VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "sql"
            VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "xml"
            VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "LNKD"
        End If
                
        Dim IncludeTablesCol As Collection
        Set IncludeTablesCol = StrSetToCol(INCLUDE_TABLES, ",")
                
        Debug.Print VCS_String.VCS_PadRight("Exporting " & obj_type_label & "...", 24);
                
        For Each td In tds
            If Len(td.connect) = 0 Then ' this is not an external table
                ' This is not a system table
                ' this is not a temporary table
                If Left$(td.name, 4) <> "MSys" And _
                   Left$(td.name, 1) <> "~" Then
                    If doTableDefs Then
                        VCS_Table.VCS_ExportTableDef td.name, obj_path
                    End If
                End If
                If INCLUDE_TABLES = "*" Then
                    DoEvents
                    ' This is not a system table
                    ' this is not a temporary table
                    If Left$(td.name, 4) <> "MSys" And _
                       Left$(td.name, 1) <> "~" Then
                        If doTableDefs Then
                            VCS_Table.VCS_ExportTableDef td.name, obj_path
                        End If

                        If doTableData Then
                            VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                        End If
                        If Len(Dir$(source_path & "tables\" & td.name & ".txt")) > 0 Then
                            obj_data_count = obj_data_count + 1
                        End If
                    End If
                
            ElseIf (Len(Replace(INCLUDE_TABLES, " ", vbNullString)) > 0) And INCLUDE_TABLES <> "*" Then
                DoEvents
                On Error GoTo Err_TableNotFound
                If InCollection(IncludeTablesCol, td.name) Then
                    If doTableData Then
                        VCS_Table.VCS_ExportTableData CStr(td.name), source_path & "tables\"
                        obj_data_count = obj_data_count + 1
                    End If
                End If
Err_TableNotFound:
                                                
                'else don't export table data
                End If
            Else
                If doTableDefs Then
                    VCS_Table.VCS_ExportLinkedTable td.name, obj_path
                End If
            End If
                                
            obj_count = obj_count + 1
        Next
        Debug.Print "[" & obj_count & "]"
        If obj_data_count > 0 Then
            Debug.Print VCS_String.VCS_PadRight("Exported data...", 24) & "[" & obj_data_count & "]"
        End If
                
                
        Debug.Print VCS_String.VCS_PadRight("Exporting Relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        VCS_Dir.VCS_MkDirIfNotExist Left$(obj_path, InStrRev(obj_path, "\"))

        VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "txt"

        Dim aRelation As DAO.Relation
                
        For Each aRelation In CurrentDb.Relations
            ' Exclude relations from system tables and inherited (linked) relations
            ' Skip if dbRelationDontEnforce property is not set. The relationship is already in the table xml file. - sean
            If Not (aRelation.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                    Or aRelation.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                    Or (aRelation.Attributes And DAO.RelationAttributeEnum.dbRelationInherited) = _
                    DAO.RelationAttributeEnum.dbRelationInherited) _
               And (aRelation.Attributes = DAO.RelationAttributeEnum.dbRelationDontEnforce) Then
                VCS_Relation.VCS_ExportRelation aRelation, obj_path & aRelation.name & ".txt"
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean

    LoadCustomisations
    
    Set Db = CurrentDb
    
    LoadCustomisations
    
    CloseFormsReports
    'InitVCS_UsingUcs2

    source_path = VCS_Dir.VCS_ProjectPath() & "source\"
    VCS_Dir.VCS_MkDirIfNotExist source_path

    Debug.Print

    If ExportQueries Then
        obj_path = source_path & "queries\"
        VCS_Dir.VCS_ClearTextFilesFromDir obj_path, "bas"
        Debug.Print VCS_String.VCS_PadRight("Exporting queries...", 24);
        SysCmd acSysCmdInitMeter, "Exporting queries", Db.QueryDefs.Count + 1
        obj_count = 0
        For Each qry In Db.QueryDefs
            DoEvents
            If Left$(qry.name, 1) <> "~" Then
                If HandleQueriesAsSQL Then
                    VCS_Query.ExportQueryAsSQL qry, obj_path & qry.name & ".bas", False
                Else
                    VCS_IE_Functions.VCS_ExportObject acQuery, qry.name, obj_path & qry.name & ".bas", VCS_File.VCS_UsingUcs2
                End If
                obj_count = obj_count + 1
            End If
            SysCmd acSysCmdUpdateMeter, obj_count
        Next
        Debug.Print VCS_String.VCS_PadRight("Sanitizing...", 15);
        VCS_IE_Functions.VCS_SanitizeTextFiles obj_path, "bas"
        Debug.Print "[" & obj_count & "]"
        SysCmd acSysCmdRemoveMeter
    End If

    
    For Each obj_type In Split( _
                               "forms|Forms|" & acForm & "," & _
                               "reports|Reports|" & acReport & "," & _
                               "macros|Scripts|" & acMacro & "," & _
                               "modules|Modules|" & acModule _
                               , "," _
                               )
        ExportObjTypeSource obj_type
    Next
    
    '-------------------------table export------------------------
    If ExportTables Then
        ExportAllTables
    End If
        
    Debug.Print "Done."
End Sub

Public Function ImportObjType(ByVal fileName As String, ByVal obj_type_label As String, ByVal obj_type_num As Integer, Optional ByVal ignoreVCS As Boolean = False, Optional ByVal src_path As String) As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    ImportObjType = 0
    obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
    obj_path = src_path & obj_type_label & "\"
    If obj_type_label = "modules" Then
        ucs2 = False
    Else
        ucs2 = VCS_File.VCS_UsingUcs2
    End If
    If IsNotVCS(obj_name) Then
        VCS_IE_Functions.VCS_ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
        ImportObjType = 1
    Else
        If ArchiveMyself And Not ignoreVCS Then
            MsgBox "Module " & obj_name & " could not be updated while running. Ensure latest version is included!", vbExclamation, "Warning"
        End If
    End If
End Function

Public Sub ImportObjTypeSource(ByVal obj_type As Variant, Optional ByVal ignoreVCS As Boolean = False, Optional ByVal src_path As String)
    Dim Db As Object ' DAO.Database
    Dim ucs2 As Boolean
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_path As String
    Dim obj_name As String
    Dim obj_count As Integer
    Dim fileName As String

    LoadCustomisations
    
    CloseFormsReports
    'InitVCS_UsingUcs2

    Set Db = CurrentDb
    
    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    obj_type_split = Split(obj_type, "|")
    obj_type_label = obj_type_split(0)
    obj_type_num = Val(obj_type_split(1))
    obj_path = src_path & obj_type_label & "\"

    If (obj_type_label = "modules") Then
        If Not VCS_Reference.VCS_ImportReferences(src_path) Then
            Debug.Print "Info: no references file in " & src_path
            Debug.Print
        End If
    End If
    
    fileName = Dir$(obj_path & "*.bas")
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing " & obj_type_label & "...", 24)
        SysCmd acSysCmdInitMeter, "Importing " & obj_type_label, 100
        obj_count = 0
        Do Until Len(fileName) = 0
            ' DoEvents no good idea!
            obj_count = obj_count + ImportObjType(fileName, obj_type_label, obj_type_num, ignoreVCS, src_path)
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
        
    End If

End Sub

Public Sub ImportAllModules(Optional ByVal ignoreVCS As Boolean = False)
    ImportObjTypeSource "modules|" & acModule, ignoreVCS
End Sub

Public Sub ImportAllForms(Optional ByVal ignoreVCS As Boolean = False, Optional ByVal src_path As String)
    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    ImportObjTypeSource "forms|" & acForm, ignoreVCS, src_path
End Sub

Public Sub ImportAllReports(Optional ByVal ignoreVCS As Boolean = False)
    ImportObjTypeSource "reports|" & acReport, ignoreVCS
End Sub

Public Sub ImportAllMacros(Optional ByVal ignoreVCS As Boolean = False)
    ImportObjTypeSource "macros|" & acMacro, ignoreVCS
End Sub

Public Sub ImportTableDef(ByVal fileName As String, Optional ByVal src_path As String)
    Dim obj_name As String
    Dim obj_path As String

    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    obj_path = src_path & "tbldef\"
    obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
    If DebugOutput Then
        Debug.Print "  [debug] table " & obj_name;
        Debug.Print
    End If
    VCS_Table.VCS_ImportTableDef CStr(obj_name), obj_path
End Sub

Public Sub ImportAllTableDefs(Optional ByVal src_path As String)
    Dim obj_path As String
    Dim fileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    obj_path = src_path & "tbldef\"
    fileName = Dir$(obj_path & "*.xml")
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing tabledefs...", 24);
        SysCmd acSysCmdInitMeter, "Importing tabledefs", 100
        obj_count = 0
        Do Until Len(fileName) = 0
            If DebugOutput Then
                If obj_count = 0 Then
                    Debug.Print
                End If
            End If
            ImportTableDef fileName, src_path
            obj_count = obj_count + 1
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If

    ' restore linked tables - we must have access to the remote store to import these!
    fileName = Dir$(obj_path & "*.LNKD")
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing Linked tabledefs...", 24);
        SysCmd acSysCmdInitMeter, "Importing Linked tabledefs", 100
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            If DebugOutput Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            VCS_Table.VCS_ImportLinkedTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

Public Sub ImportTableData(ByVal fileName As String, Optional ByVal src_path)
    Dim appendOnly As Boolean
    Dim obj_name As String
    Dim obj_path As String
    
    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    obj_path = src_path & "tables\"
    appendOnly = False
    obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
    If InStrRev(obj_name, ".") Then
        ' For now assume it is append if the extra . exists.
        obj_name = Mid$(obj_name, 1, InStrRev(obj_name, ".") - 1)
        appendOnly = True
    End If
            
    VCS_Table.VCS_ImportTableData CStr(obj_name), obj_path & fileName, appendOnly
End Sub

Public Sub ImportAllTableData(Optional ByVal src_path As String)
    Dim obj_path As String
    Dim fileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    If src_path = "" Then
        src_path = VCS_SourcePath
    End If
    obj_path = src_path & "tables\"
    fileName = Dir$(obj_path & "*.xml")
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing tables...", 24);
        SysCmd acSysCmdInitMeter, "Importing tables", 100
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            ImportTableData fileName, src_path
            obj_count = obj_count + 1
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

Public Sub ImportAllTableDataMacros()
    Dim obj_path As String
    Dim fileName As String
    Dim obj_count As Integer
    Dim obj_name As String

    obj_path = VCS_SourcePath & "tbldef\"
    fileName = Dir$(obj_path & "*.dm")
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing Data Macros...", 24);
        SysCmd acSysCmdInitMeter, "Importing Data Macros", 100
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            'VCS_Table.VCS_ImportTableData CStr(obj_name), obj_path
            VCS_DataMacro.VCS_ImportDataMacros obj_name, obj_path
            obj_count = obj_count + 1
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
End Sub

Public Function VCS_SourcePath() As String
    VCS_SourcePath = VCS_Dir.VCS_ProjectPath() & "source\"
End Function

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource(Optional ByVal ignoreVCS As Boolean = False)
    Dim FSO As Object
    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean
    Dim appendOnly As Boolean

    Set FSO = CreateObject("Scripting.FileSystemObject")

    LoadCustomisations
    
    CloseFormsReports
    'InitVCS_UsingUcs2

    source_path = VCS_SourcePath
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print
    

    obj_path = source_path & "queries\"
    fileName = Dir$(obj_path & "*.bas")
    
    Dim tempFilePath As String
    tempFilePath = VCS_File.VCS_TempFile()
    
    If Len(fileName) > 0 Then
        Debug.Print VCS_String.VCS_PadRight("Importing queries...", 24);
        SysCmd acSysCmdInitMeter, "Importing queries", 100
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
            'Check for plain sql export/import
            if HandleQueriesAsSQL then
                VCS_Query.ImportQueryFromSQL obj_name, obj_path & fileName, False
            Else
                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, obj_path & fileName, VCS_File.VCS_UsingUcs2
                VCS_IE_Functions.VCS_ExportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2
                VCS_IE_Functions.VCS_ImportObject acQuery, obj_name, tempFilePath, VCS_File.VCS_UsingUcs2
            End if          
            obj_count = obj_count + 1
            fileName = Dir$()
            SysCmd acSysCmdUpdateMeter, obj_count
        Loop
        SysCmd acSysCmdRemoveMeter
        Debug.Print "[" & obj_count & "]"
    End If
    
    VCS_Dir.VCS_DelIfExist tempFilePath

    ' restore table definitions
    ImportAllTableDefs
    
    ' NOW we may load data
    ImportAllTableData
    
    'load Data Macros - not DRY!
    ImportAllTableDataMacros

    'import Data Macros
    

    For Each obj_type In Split( _
                               "forms|" & acForm & "," & _
                               "reports|" & acReport & "," & _
                               "macros|" & acMacro & "," & _
                               "modules|" & acModule _
                               , "," _
                               )
        ImportObjTypeSource obj_type, ignoreVCS
    Next
    
    'import Print Variables
    Debug.Print VCS_String.VCS_PadRight("Importing Print Vars...", 24);
    obj_count = 0
    
    obj_path = source_path & "reports\"
    fileName = Dir$(obj_path & "*.pv")
    Do Until Len(fileName) = 0
        DoEvents
        obj_name = Mid$(fileName, 1, InStrRev(fileName, ".") - 1)
        VCS_Report.VCS_ImportPrintVars obj_name, obj_path & fileName
        obj_count = obj_count + 1
        fileName = Dir$()
    Loop
    Debug.Print "[" & obj_count & "]"
    
    'import relations
        Debug.Print VCS_String.VCS_PadRight("Importing Relations...", 24);
        obj_count = 0
        obj_path = source_path & "relations\"
        fileName = Dir$(obj_path & "*.txt")
        Do Until Len(fileName) = 0
            DoEvents
            VCS_Relation.VCS_ImportRelation obj_path & fileName
            obj_count = obj_count + 1
            fileName = Dir$()
        Loop
        Debug.Print "[" & obj_count & "]"
    DoEvents
    
    Debug.Print "Done."
End Sub

' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject()
    On Error GoTo ErrorHandler

    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr$(149) & " Tables" & vbCrLf & _
              Chr$(149) & " Forms" & vbCrLf & _
              Chr$(149) & " Macros" & vbCrLf & _
              Chr$(149) & " Modules" & vbCrLf & _
              Chr$(149) & " Queries" & vbCrLf & _
              Chr$(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Sub
    End If

    Dim Db As DAO.Database
    Set Db = CurrentDb
    CloseFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print
    
    Dim rel As DAO.Relation
    For Each rel In CurrentDb.Relations
        If Not (rel.name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or _
                rel.name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
            CurrentDb.Relations.Delete (rel.name)
        End If
    Next
        
        ' First gather all Query Names. 
        ' If you delete right away, the iterator loses track and only deletes every 2nd Query
        Dim toBeDeleted As Collection
    Set toBeDeleted = New Collection
    Dim qryName As Variant
    
    Debug.Print "Deleting queries"
    Dim dbObject As Object
    For Each dbObject In Db.QueryDefs
        DoEvents
        If Left$(dbObject.name, 1) <> "~" Then
            toBeDeleted.Add dbObject.Name
        End If
    Next
        
        For Each qryName In toBeDeleted
        Db.QueryDefs.Delete qryName
    Next
        
        Set toBeDeleted = Nothing
    
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If Left$(td.name, 4) <> "MSys" And _
           Left$(td.name, 1) <> "~" Then
            CurrentDb.TableDefs.Delete (td.name)
        End If
    Next

    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME As Byte = 0
    Const OTID As Byte = 1

    For Each objType In Split( _
                              "Forms|" & acForm & "," & _
                              "Reports|" & acReport & "," & _
                              "Scripts|" & acMacro & "," & _
                              "Modules|" & acModule _
                              , "," _
                              )
        objTypeArray = Split(objType, "|")
        DoEvents
        For Each doc In Db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left$(doc.name, 1) <> "~") And _
               (IsNotVCS(doc.name)) Then
                '                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.name
            End If
        Next
    Next
    
    Debug.Print "================="
    Debug.Print "Importing Project"
    ImportAllSource
    
    Exit Sub

ErrorHandler:
    Debug.Print "VCS_ImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
         Err.Description
End Sub


'===================================================================================================================================
'-----------------------------------------------------------'
' Helper Functions - these should be put in their own files '
'-----------------------------------------------------------'

' Close all open forms.
Public Sub CloseFormsReports()
    On Error GoTo ErrorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).Name
        DoEvents
    Loop
    Exit Sub

ErrorHandler:
    Debug.Print "VCS_ImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & _
         Err.Description
End Sub


'errno 457 - duplicate key (& item)
Private Function StrSetToCol(ByVal strSet As String, ByVal delimiter As String) As Collection 'throws errors
    Dim strSetArray() As String
    Dim col As Collection
    
    Set col = New Collection
    strSetArray = Split(strSet, delimiter)
    
    Dim strPart As Variant
    For Each strPart In strSetArray
        col.Add strPart, strPart
    Next
    
    Set StrSetToCol = col
End Function


' Check if an item or key is in a collection
Private Function InCollection(col As Collection, Optional vItem, Optional vKey) As Boolean
    On Error Resume Next

    Dim vColItem As Variant

    InCollection = False

    If Not IsMissing(vKey) Then
        col.Item vKey

        '5 if not in collection, it is 91 if no collection exists
        If Err.Number <> 5 And Err.Number <> 91 Then
            InCollection = True
        End If
    ElseIf Not IsMissing(vItem) Then
        For Each vColItem In col
            If vColItem = vItem Then
                InCollection = True
                GoTo Exit_Proc
            End If
        Next vColItem
    End If

Exit_Proc:
    Exit Function
Err_Handle:
    Resume Exit_Proc
End Function

Public Function getImages(control As Object, ByRef image)
    Dim frmRibbonImages As Form ' USysRibbonImages
    Dim rsForm As DAO.Recordset2
    
    On Error Resume Next
    If frmRibbonImages Is Nothing Then
        DoCmd.OpenForm "USysRibbonImages", WindowMode:=acHidden
        Set frmRibbonImages = Forms("USysRibbonImages")
    End If
    Set rsForm = frmRibbonImages.Recordset
    
    rsForm.FindFirst "ControlID='" & control.ID & "'"
    If rsForm.NoMatch Then
        ' No image found
        Set image = Nothing
    Else
        Set image = frmRibbonImages.Images.PictureDisp
    End If
End Function

Public Sub ShowTablesWithData()
    Dim tbl As TableDef
    For Each tbl In CurrentDb.TableDefs
      If Left(tbl.name, 4) <> "mSys" And tbl.RecordCount > 0 Then
        Debug.Print tbl.name, tbl.RecordCount
      End If
    Next tbl
End Sub
