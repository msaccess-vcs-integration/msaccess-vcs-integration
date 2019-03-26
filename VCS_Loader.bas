Attribute VB_Name = "VCS_Loader"
Option Compare Database

Option Explicit

Public Sub loadVCS(Optional ByVal SourceDirectory As String)
    If SourceDirectory = vbNullString Then
        SourceDirectory = CurrentProject.Path & "\MSAccess-VCS\"
        On Error Resume Next
        If Not ((GetAttr(SourceDirectory) And vbDirectory) <> vbDirectory) Then
            ' Ask the user where the source is.
            SourceDirectory = VCS_GetFolder
            ' InputBox("Please specify the directory where the VCS files exist:", "VCS Load", CurrentProject.Path)
        End If
    End If

'check if directory exists! - SourceDirectory could be a file or not exist
On Error GoTo Err_DirCheck
    If ((GetAttr(SourceDirectory) And vbDirectory) = vbDirectory) Then
        GoTo Fin_DirCheck
    Else
        'SourceDirectory is not a directory
        Err.Raise 60000, "loadVCS", "Source Directory specified is not a directory"
    End If

Err_DirCheck:
    
    If Err.Number = 53 Then 'SourceDirectory does not exist
        Debug.Print Err.Number & " | " & "File/Directory not found"
    Else
        Debug.Print Err.Number & " | " & Err.Description
    End If
    Exit Sub
Fin_DirCheck:

    'delete if modules already exist + provide warning of deletion?

    On Error GoTo Err_DelHandler

    Dim fileName As String
    'Use the list of files to import as the list to delete
    fileName = Dir$(SourceDirectory & "VCS_*.bas")
    Do Until Len(fileName) = 0
        'strip file type from file name
        fileName = Left$(fileName, InStrRev(fileName, ".bas") - 1)
        If Not fileName = "VCS_Loader" Then
            DoCmd.DeleteObject acModule, fileName
        End If
        fileName = Dir$()
    Loop

    GoTo Fin_DelHandler
    
Err_DelHandler:
    If Err.Number <> 7874 Then 'is not - can't find object
        Debug.Print "WARNING (" & Err.Number & ") | " & Err.Description
    End If
    Resume Next
    
Fin_DelHandler:
    fileName = vbNullString

'import files from specific dir? or allow user to input their own dir?
On Error GoTo Err_LoadHandler

    fileName = Dir$(SourceDirectory & "VCS_*.bas")
    Do Until Len(fileName) = 0
        'strip file type from file name
        fileName = Left$(fileName, InStrRev(fileName, ".bas") - 1)
        If Not fileName = "VCS_Loader" Then
            Application.LoadFromText acModule, fileName, SourceDirectory & fileName & ".bas"
        End If
        fileName = Dir$()
    Loop

    GoTo Fin_LoadHandler
    
Err_LoadHandler:
    Debug.Print Err.Number & " | " & Err.Description
    Resume Next

Fin_LoadHandler:
    VCS_Bootstrap SourceDirectory
    Debug.Print "Done"
    
    displayFormVersion
End Sub

Public Sub displayFormVersion()
    Dim versionPath As String, FormsVersion As String, textline As String, posLat As Integer, posLong As Integer
    versionPath = CurrentProject.Path & "\VERSION.txt"
    Open versionPath For Input As #1

    Do Until EOF(1)
        Line Input #1, textline
        FormsVersion = FormsVersion & textline
        
    Loop
    Close #1

    MsgBox "Form Version: " & FormsVersion & " loaded"
End Sub

Public Sub VCS_ImportAllSource()
    VCS_LoadAllSource
End Sub

Public Sub VCS_ImportAllModules()
    VCS_LoadModules
End Sub

Public Sub VCS_LoadStub()
    loadVCS
End Sub

Public Sub VCS_LoadModules()
    LoadCustomisations
    ImportAllModules True
End Sub

Public Sub VCS_LoadAllSource()
    LoadCustomisations
    ImportAllSource True
End Sub

Private Sub VCS_Bootstrap(ByVal SourceDirectory As String)
    VCS_Reference.VCS_ImportReferences(SourceDirectory)
    LoadCustomisations
    VCS_Bootstrap_Tables SourceDirectory
End Sub

Private Function VCS_GetFolder() As String
    Const msoFileDialogFolderPicker As Long = 4
    Dim fldr As Object
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = CurrentProject.Path
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    VCS_GetFolder = sItem & "\"
    Set fldr = Nothing
End Function

Private Sub VCS_Bootstrap_Tables(ByVal SourceDirectory As String)
    CloseFormsReports
    ImportAllTableDefs(SourceDirectory)
    ImportAllTableData(SourceDirectory)
    ImportAllForms ignoreVCS:=False, src_path:=SourceDirectory
End Sub
