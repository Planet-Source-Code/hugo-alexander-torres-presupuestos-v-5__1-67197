Attribute VB_Name = "modGlobals"
Option Explicit

'**********************************************************************************
'FindFirstFile, FindNextFile, FindClose
'**********************************************************************************
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Public Enum FILE_ATTRIBUTES
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

'**********************************************************************************
'GetTickCount, Used for timing events
'**********************************************************************************
Declare Function GetTickCount Lib "kernel32" () As Long

Public Function CompactDatabase(strDatabaseName As String) As Boolean
        On Error GoTo Err_CompactDatabase
        Dim strPath As String
        Dim strPath1 As String
        Dim strPathSize As String
        Dim strPathSize2 As String
        Dim time1
        
        Screen.MousePointer = vbHourglass
        
        time1 = Time
        'Save Paths for Database
        strPath = strDatabaseName
        strPath1 = Left(strDatabaseName, Len(strDatabaseName) - 4) & "Backup.mdb"
        'Get Size of File Before Compacting
        strPathSize = GetFileSize(strPath)
        'Kill the file if it exists
        If Dir(strPath1) <> "" Then Kill strPath1
        'Compact Database to New Name
        DBEngine.CompactDatabase strPath, strPath1
        ''Kill the file if it exists
        If Dir(strPath) <> "" Then Kill strPath
        'Compact back to original Name
        DBEngine.CompactDatabase strPath1, strPath
        'Kill the file, no need to save it
        If Dir(strPath1) <> "" Then Kill strPath1
        'Get Size of File After Compacting
        strPathSize2 = GetFileSize(strPath)
        CompactDatabase = True
        
        
        'Display the Summary
        
        MsgBox UCase(strDatabaseName) & " compactada exitósamente. " _
        & vbNewLine & vbNewLine & "Tiempo empleado:" & vbTab & vbTab & vbTab & Hour(Time - time1) & "horas   " & Minute(Time - time1) & "minutos   " & Second(Time - time1) & "segundos " _
        & vbNewLine & "Tamaño antes de compactar  :" & vbTab & strPathSize _
        & vbNewLine & "Tamaño después de compactar:" & vbTab & strPathSize2, vbInformation, "Presup Ver. 3.0  -  Reparación Exitosa"
       
Err_CompactDatabase:
    
    
        Select Case Err
            Case 0
            Case Else
            MsgBox Err & ": " & Error, vbCritical, "Error de compactación"
        End Select
    
    Screen.MousePointer = vbNormal
End Function

Public Function GetFileSize(strFile As String) As String
Dim fso As New Scripting.FileSystemObject
Dim f As File
Dim lngBytes As Long
Const KB As Long = 1024
Const MB As Long = 1024 * KB
Const GB As Long = 1024 * MB
    
    Set f = fso.GetFile(fso.GetFile(strFile))
    lngBytes = f.Size


    If lngBytes < KB Then
        GetFileSize = Format(lngBytes) & " bytes"
    ElseIf lngBytes < MB Then
        GetFileSize = Format(lngBytes / KB, "0.00") & " KB"
    ElseIf lngBytes < GB Then
        GetFileSize = Format(lngBytes / MB, "0.00") & " MB"
    Else
        GetFileSize = Format(lngBytes / GB, "0.00") & " GB"
    End If
    
End Function

Public Function RemoveFileFromPath(strFilePath As String) As Boolean
On Error GoTo ErrPath

    Kill strFilePath
    
Success:
    RemoveFileFromPath = True
    Exit Function
ErrPath:
    RemoveFileFromPath = False

End Function
