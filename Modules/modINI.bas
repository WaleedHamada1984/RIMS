Attribute VB_Name = "modINI"
'*****************************************************************************
'DLL declarations
'*****************************************************************************
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Global Provider_DB2 As String
Global Provider_SQL As String
Global Server_DB2 As String
Global User_DB2 As String
Global PWD_DB2 As String
Global MyIP As String
Global Server_SQL As String
Global User_SQL As String
Global PWD_SQL As String
Global DB_SQL As String
Global Pub_DataLib As String

Public Function ReadIniFile(Fname As String) As Boolean
Dim ChkExist As String
On Error GoTo ErrorHandler
ReadIniFile = False

If Right(App.Path, 1) = "\" Then
    ChkExist = Dir(App.Path & Fname)
Else
    ChkExist = Dir(App.Path & "\" & Fname)
End If

If ChkExist = "" Then
    MsgBox "The following file - " & UCase(Fname) & " - does not exist in the current directory. Unable to continue.", vbCritical, "Error"
    Exit Function
End If

Fname = App.Path & "\" & Fname

Provider_DB2 = GetFromINI("Database", "ProviderDB2", "", Fname)
Provider_SQL = GetFromINI("Database", "ProviderSQL", "", Fname)
User_DB2 = GetFromINI("Database", "UserDB2", "", Fname)
PWD_DB2 = GetFromINI("Database", "PWDDB2", "", Fname)
Server_DB2 = GetFromINI("Database", "ServerDB2", "", Fname)
User_SQL = GetFromINI("Database", "UserSQL", "", Fname)
PWD_SQL = GetFromINI("Database", "PWDSQL", "", Fname)
Server_SQL = GetFromINI("Database", "ServerSQL", "", Fname)
DB_SQL = GetFromINI("Database", "DBSQL", "", Fname)
Pub_DataLib = GetFromINI("Database", "DATALIB", "", Fname)
Pub_CtlLib = GetFromINI("Database", "CTLLIB", "", Fname)
ReadIniFile = True

Exit Function

ErrorHandler:
    MsgBox "Unexpected error " & Err.Description, vbCritical, "Error"

End Function
'// Functions
Function GetFromINI(sSection As String, sKey As String, sDefault As String, sIniFile As String)
    Dim sBuffer As String, lRet As Long
    sBuffer = String$(255, 0)
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        If sDefault <> "" Then AddToINI sSection, sKey, sDefault, sIniFile
        GetFromINI = sDefault
    Else
        GetFromINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    End If
End Function
'// Returns True if successful. If section does not
'// exist it creates it.
Function AddToINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    AddToINI = (lRet)
End Function

