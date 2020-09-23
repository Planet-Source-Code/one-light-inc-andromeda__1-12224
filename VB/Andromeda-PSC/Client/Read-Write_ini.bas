Attribute VB_Name = "INI"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

Function ReadINI(AppName$, Keyname$, filename$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   ReadINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal Keyname$, "", RetStr, Len(RetStr), filename$))
End Function

Function WriteINI(mizainz$, Place$, Toput$, AppName$)
    r% = WritePrivateProfileString(mizainz$, Place$, Toput$, AppName$)
End Function

