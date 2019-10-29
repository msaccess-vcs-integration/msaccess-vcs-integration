Attribute VB_Name = "VCS_Query"
Option Compare Database
Option Explicit

Const StartConnect As String = "[CONNECT]"
Const StartSQL As String = "[SQL]"
Const StartRecRecords As String = "[ReturnRecs]"

Public Sub ExportQueryAsSQL(qry As QueryDef, ByVal file_path As String, _
                            Optional ByVal Ucs2Convert As Boolean = False)

    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        writeTextToFile qry.sql, tempFileName
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, file_path
    Else
        If Not (qry.connect = "") Then
            Dim qryconnect As String
            qryconnect = qry.connect
            If (Right(qryconnect, 1) = vbLf) Then
                qryconnect = Left(qryconnect, Len(qryconnect) - 2)
            End If
            writeTextToFile StartConnect & qryconnect & vbCrLf & StartRecRecords & qry.ReturnsRecords & vbCrLf & StartSQL & vbCrLf & qry.sql, file_path
        Else
            writeTextToFile qry.sql, file_path
        End If
    End If

End Sub

Private Sub writeTextToFile(ByVal textToWrite As String, ByVal file_path As String)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(file_path)

    oFile.WriteLine textToWrite
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Sub

Private Function readFromTextFile(ByVal file_path As String) As String
    Dim textRead As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.OpenTextFile(file_path, ForReading)

    Do While Not oFile.AtEndOfStream
        textRead = textRead & oFile.ReadLine & vbCrLf
    Loop

    readFromTextFile = textRead
    
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Function

Public Sub ImportQueryFromSQL(ByVal obj_name As String, ByVal file_path As String, _
                                Optional ByVal Ucs2Convert As Boolean = False)
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    Set db = CurrentDb
    
    If Ucs2Convert Then
        
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 file_path, tempFileName
        On Error Resume Next
        db.QueryDefs.Delete (obj_name)
        db.CreateQueryDef obj_name, readFromTextFile(file_path)
        
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile tempFileName
    Else
        Dim fileStr As String
        Dim ConnectString As String
        Dim SQLString As String
        Dim SQLPosition As Integer
        Dim retrecpos As Integer
        Dim retrecstr As String
        Dim retrec As Boolean
        'normally, queries are going to reurn records, so set to true in case it's not defined.
        retrec = True

        fileStr = readFromTextFile(file_path)
        'find out if there's a connect string or other stuff. If there is, the SQL position will be greater than 0
        SQLPosition = InStr(1, fileStr, StartSQL, vbBinaryCompare)
        
        If SQLPosition > 0 Then
            retrecpos = InStr(1, fileStr, StartRecRecords, vbBinaryCompare)
            retrecstr = mid(fileStr, retrecpos + Len(StartRecRecords), 1)
            ConnectString = mid(fileStr, 10, retrecpos - 11)
            'find the start of the SQL, plus two charachters (carriage return + line feed = vbcrlf)
            SQLString = mid(fileStr, SQLPosition + Len(StartSQL) + 2)

            Select Case retrecstr
                Case "f", "F", "0"
                    retrec = False
                Case "T", "t", "1"
                    retrec = True
            End Select
            
            On Error Resume Next
            DoCmd.DeleteObject acQuery, obj_name
            Set qdf = db.CreateQueryDef(obj_name)
            Set qdf = db.QueryDefs(obj_name)
            With qdf
                .connect = ConnectString
                .ReturnsRecords = retrec
                .SQL = SQLString
            End With
        Else
            On Error Resume Next
            DoCmd.DeleteObject acQuery, obj_name
            Set qdf = db.CreateQueryDef(obj_name, fileStr)
        End If

    End If
    qdf.Close
    Set qdf = Nothing
    Set db = Nothing
End Sub
