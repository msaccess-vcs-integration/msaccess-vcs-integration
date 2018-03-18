Attribute VB_Name = "VCS_IniHandler"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function GetSectionEntry(sSection As String, _
                         sKey As String, _
                         sIniFile As String) As String

    On Error GoTo Err_GetSectionEntry

    ' Variable declarations.
    Dim sRetBuf         As String
    Dim iLenBuf         As Integer
    Dim sFileName       As String
    Dim sReturnValue    As String
    Dim lRetVal         As Long
    
    ' Set the return buffer to by 256 spaces. This should be enough to
    ' hold the value being returned from the INI file, but if not,
    ' increase the value.
    sRetBuf = Space(256)

    ' Get the size of the return buffer.
    iLenBuf = Len(sRetBuf)

    ' Read the INI Section/Key value into the return variable.
    sReturnValue = GetPrivateProfileString(sSection, _
                                           sKey, _
                                           "", _
                                           sRetBuf, _
                                           iLenBuf, _
                                           sIniFile)

    ' Trim the excess garbage that comes through with the variable.
    sReturnValue = Trim(Left(sRetBuf, sReturnValue))

    ' If we get a value returned, pass it back as the argument.
    ' Else pass "False".
    If Len(sReturnValue) > 0 Then
        GetSectionEntry = sReturnValue
    Else
        GetSectionEntry = vbNullString
    End If
    
Exit_Clean:
    Exit Function
    
Err_GetSectionEntry:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Clean

End Function

Function SetSectionEntry(sSection As String, _
                         sKey As String, _
                         sValue As String, _
                         sIniFile As String) As Long

    On Error GoTo Err_SetSectionEntry

    ' Variable declarations.
    Dim lRetVal         As Long
    If sValue = "" Then
        ' Force delete of key.
        sValue = vbNullString
    End If
    
    
    
    ' Write to the INI file and capture the value returned
    ' in the API function.
    lRetVal = WritePrivateProfileString(sSection, _
                                        sKey, _
                                        sValue, _
                                        sIniFile)

    ' Check to see if we had an error wrting to the INI file.
    If lRetVal = 0 Then
        SetSectionEntry = 1
    Else
        SetSectionEntry = 0
    End If

Exit_Clean_2:
    Exit Function
    
Err_SetSectionEntry:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Clean_2

End Function

Function GetSectionEntryArray(sSection As String, _
                              sKey As String, _
                              sIniFile As String) As String()

    Dim sValue As String
    Dim aValue() As String
    Dim index As Integer
    
    sValue = GetSectionEntry(sSection, sKey, sIniFile)
    aValue = Split(sValue, ",")
    For index = LBound(aValue) To UBound(aValue)
        aValue(index) = Trim(aValue(index))
    Next index
    GetSectionEntryArray = aValue
    
End Function

