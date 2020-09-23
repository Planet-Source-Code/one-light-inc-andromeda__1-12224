Attribute VB_Name = "EplorerMod1"
Public ClientPath As String
Public ServerPath As String


Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFilename As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
       (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
       (ByVal lpRootPathName As String, _
       lpSectorsPerCluster As Long, _
       lpBytesPerSector As Long, _
       lpNumberOfFreeClusters As Long, _
       lpTotalNumberOfClusters As Long) As Long
   

Public NbFile As Long
Public FileFSToOpen As String
Public StringToFind As String
Public ProgressCancel As Boolean
Public TypeView

Public Const MAX_PATH As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
   
Type SaveF
   StingToSave As String
End Type
   
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Function RenameFile(xPath As String, xNewName As String) As Boolean
    'Renames a disk file
    If Exists(xPath) = False Then RenameFile = False: Exit Function
    
    Dim Fiz As File
    Dim fizile As FileSystemObject
    Set fizile = CreateObject("Scripting.FileSystemObject")
    
    Set Fiz = fizile.GetFile(xPath)
    
    Fiz.Name = xNewName
    
    RenameFile = True
    
    Set Fiz = Nothing
    Set fizile = Nothing
    
End Function
        

Function RenameFolder(xPath As String, xNewName As String) As Boolean
    'Renames a folder
    
    Dim Fld As Folder
    Dim fizile As FileSystemObject
    Set fizile = CreateObject("Scripting.FileSystemObject")
    
    If fizile.FolderExists(xPath) = False Then RenameFolder = False: Exit Function
    
    Set Fld = fizile.GetFolder(xPath)
    
    Fld.Name = xNewName
    
    RenameFolder = True
    
    Set Fld = Nothing
    Set fizile = Nothing
    
End Function
Sub ChowFromFolder(FileList As ListView, ByVal zpath As String, ByVal FileType As String)
'this function was found on planet-source-code.com, but has been slightly modified

On Error Resume Next
    
       If Right(zpath, 1) <> "\" Then
            zpath = zpath + "\"
       End If
       
       ClientPath = zpath
       frmFileView.Status.Panels(1).Text = UCase(ClientPath)
       
       FileList.ListItems.Clear
       
       Dim hFile As Long, result As Long, szPath As String
       Dim WFD As WIN32_FIND_DATA
       Dim TMP As ListItem
       Dim pos1
       
       szPath = zpath & FileType & Chr$(0)
       'Start asking windows for files.
       hFile = FindFirstFile(szPath, WFD)
       Do
           ts = StripNull(WFD.cFileName)
           If Not (ts = "." Or ts = "..") Then
             
             If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                rep = "folder"
                rep2 = " Directory"
             Else
                ext$ = objFso.GetExtensionName(WFD.cFileName)
                rep = GetImage(ext$)
                rep2 = GetType(ext$)
             End If
             pos1 = InStr(1, WFD.cFileName, Chr$(0), vbBinaryCompare)
             If Trim(Mid(WFD.cFileName, 1, pos1 - 1)) <> "" Then
                Set TMP = FileList.ListItems.Add(, , Trim(WFD.cFileName), rep, rep)
                If rep <> "folder" Then TMP.SubItems(1) = FormatFileSize(WFD.nFileSizeLow)
                Dim strtemp As Variant
                TMP.SubItems(2) = rep2
             End If
           End If
next1:
             WFD.cFileName = ""
             result = FindNextFile(hFile, WFD)
       Loop Until result = 0
       FindClose hFile
       
       With frmFileView.ClientDrives
            For x = 1 To .ComboItems.Count
                If LCase(.ComboItems(x).Text) = LCase(ClientPath) Then
                    .ComboItems(x).Selected = True
                    GoTo next2
                End If
            Next x
            
            Dim itm As ComboItem
            Set itm = .ComboItems.Add(, , ClientPath, "folder")
            itm.Selected = True
       End With
next2:
End Sub

Function GetImage(ByVal imgstr As String) As String
Select Case LCase(imgstr)
    Case "exe", "rtf", "txt", "ini", "dll", "zip", "doc"
        GetImage = LCase(imgstr)
    Case "bmp", "jpg", "gif"
        GetImage = "picture"
    Case "mp3", "wav"
        GetImage = "sound"
    Case "mpg", "avi"
        GetImage = "media"
    Case "fnt", "ttf", "fon"
        GetImage = "font"
    Case Else
        GetImage = "misc"
End Select
End Function

Function GetType(fileExt As String) As String

Select Case LCase(fileExt)
    Case "exe"
        GetType = "Program"
    Case "dll"
        GetType = "Application Extention"
    Case "rtf"
        GetType = "Rich Text File"
    Case "txt"
        GetType = "Text File"
    Case "doc"
        GetType = "Micro$oft Word Document"
    Case "bmp"
        GetType = "Bitmap File"
    Case "jpg"
        GetType = "JPEG"
    Case "gif"
        GetType = "GIF"
    Case "avi"
        GetType = "Movie"
    Case "mpg"
        GetType = "Media File"
    Case "mp3"
        GetType = "MPEG Audio File"
    Case "wav"
        GetType = "WAV File"
    Case "ini"
        GetType = "INI File"
    Case "fnt", "ttf", "fon"
        GetType = "Font"
    Case "zip"
        GetType = "Zip File"
    Case Else
        GetType = UCase(fileExt) + " File"
End Select
End Function



Public Function StripNull(ByVal WhatStr As String) As String
   Dim pos As Integer
    pos = InStr(WhatStr, Chr$(0))
    If pos > 0 Then
       StripNull = Left$(WhatStr, pos - 1)
    Else
       StripNull = WhatStr
    End If
End Function


           
