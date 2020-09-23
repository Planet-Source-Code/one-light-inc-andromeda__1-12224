Attribute VB_Name = "ServerMod"
'//////////////////////////////////////////////////////////
'// Module for Andromeda 1.0 Remote File Server for      //
'// Microsoft Win32 by Ryan and Andrew Lederman          //
'// www.induhviduals.com/andromeda                       //
'//////////////////////////////////////////////////////////

Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const HKEY_LOCAL_MACHINE = &H80000002

Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long


Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long


Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Const MAX_PATH& = 260

Public sEnabled As Boolean
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
    End Type

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public filesOpen As Integer
Public strBuffer As String
Public PacketCount As Integer
Public EngineRunning As Boolean
Public SendPort(1 To 100)
Public intMax2 As Integer
Public fileNum As Long
Public FileSize1 As Long
Public fileName As String
Public StartSending  As Boolean
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    'constants required by Shell_NotifyIcon API call:
    Public Const NIM_ADD = &H0
    Public Const NIM_MODIFY = &H1
    Public Const NIM_DELETE = &H2
    Public Const NIF_MESSAGE = &H1
    Public Const NIF_ICON = &H2
    Public Const NIF_TIP = &H4
    Public Const WM_MOUSEMOVE = &H200

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public nid As NOTIFYICONDATA

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Public Const REG_NONE = (0)                         'No value type
Public Const REG_SZ = (1)                           'Unicode nul terminated string
Public Const REG_EXPAND_SZ = (2)                    'Unicode nul terminated string w/enviornment var
Public Const REG_BINARY = (3)                       'Free form binary
Public Const REG_DWORD = (4)                        '32-bit number
Public Const REG_DWORD_LITTLE_ENDIAN = (4)          '32-bit number (same as REG_DWORD)
Public Const REG_DWORD_BIG_ENDIAN = (5)             '32-bit number
Public Const REG_LINK = (6)                         'Symbolic Link (unicode)
Public Const REG_MULTI_SZ = (7)                     'Multiple Unicode strings
Public Const REG_RESOURCE_LIST = (8)                'Resource list in the resource map
Public Const REG_FULL_RESOURCE_DESCRIPTOR = (9)     'Resource list in the hardware description
Public Const REG_RESOURCE_REQUIREMENTS_LIST = (10)
Const READ_CONTROL = &H20000
Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Boolean
End Type
Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
End Type
Type ACL
        AclRevision As Byte
        Sbz1 As Byte
        AclSize As Integer
        AceCount As Integer
        Sbz2 As Integer
End Type
Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal SecurityInformation As Long, pSecurityDescriptor As SECURITY_DESCRIPTOR) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long

Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, _
    ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long

Function CreateFolder(xFolder As String, Winsock As Winsock)
    'Creates a new folder. if function fails
    'returns False
    On Error GoTo failCreate
    MkDir xFolder
    sOutput "MKDIR '" & xFolder & "' from IP '" & Winsock.RemoteHostIP & "'"
    Winsock.SendData ("CREATED")
    Exit Function
failCreate:
    Winsock.SendData ("NOTCREATED")
End Function

Function DisplayLogFile(strWhichLog As String)
    'Opens the specified log file and displays to user
    Dim WhichFile As String
    With frmLog
    
    Select Case strWhichLog
        Case "Login":
            .Caption = "ndromeda - Log File (Logins)"
            WhichFile = App.Path + "\Log.txt"
            
        Case "FileTransfer":
            .Caption = "ndromeda - Log File (File Transfers)"
            WhichFile = App.Path + "\FTransfer.txt"
        Case "Output":
            .Caption = "ndromeda - Log File (Server Output)"
            WhichFile = App.Path + "\Output.txt"
    End Select
    
    If Exists(WhichFile) = False Then
        MsgBox "The Log file: " & vbCrLf & WhichFile & vbCrLf & "was not found. Andromeda will now create a new, empty log.", 16, "Error: File Not Found"
        i = FreeFile
        Open WhichFile For Output As #i
        Close #i
    End If
    
i = FreeFile
Open WhichFile For Input As #i
    Do While Not EOF(i): DoEvents
    
    Line Input #i, Record$
    
    Entire = Entire + Record$ + vbCrLf
    
    Loop
Close #i

.txtLogin.Text = Entire
.lblFileSize.Caption = FileLen(WhichFile) & " bytes"
.WhichLog.Text = strWhichLog
.Show , frmMain
End With

End Function

Sub EnableServer(WhichState As Boolean)
    'Takes a boolean, toggles server state depending on value passed
    'True = Enabled, False = Disabled
    Select Case WhichState
    
    Case True:
         frmMain.Server(0).Close
        frmMain.Server(0).LocalPort = 6969
         frmMain.Server(0).Listen
         Do While frmMain.Server(0).State <> sckListening
            DoEvents
            If frmMain.Server(0).State = sckError Then
                MsgBox "An error occurred while trying to initialize the listening socket.", 16, "Error": Exit Sub
            End If
        Loop
        sEnabled = True
        frmMain.TimerUptime.Enabled = True
        frmMain.Caption = "ndromeda RFS (Enabled)"
        sOutput "Server Enabled"
    Case False:
        frmMain.Server(0).Close
        Do While frmMain.Server(0).State <> sckClosed
            DoEvents
        Loop
        sEnabled = False
        frmMain.TimerUptime.Enabled = False
        frmMain.txtElapsed.Caption = "00:00:00"
        frmMain.Caption = "ndromeda RFS (Disabled)"
        sOutput "Server Disabled"
    End Select
    
End Sub

Function InvalidMessage() As String
    'Reads the Invalid Message file (\imessage.txt)
    'and returns the contents
    Dim fileNum As Integer
    fileNum = FreeFile
    If Exists(App.Path + "\imessage.txt") = False Then
        Open App.Path + "\imessage.txt" For Output As #fileNum
        Close #fileNum
        InvalidMessage = ""
    Exit Function
    End If
    
    Open App.Path + "\imessage.txt" For Input As #fileNum
        Do While Not EOF(fileNum): DoEvents
        
        Line Input #fileNum, Record$
        
        Entire = Entire + Record$ + vbCrLf
        Loop
    Close #fileNum
    InvalidMessage = Entire
End Function

Function IsValidSharedFolder(strFolder As String) As Boolean
    'Takes the path of a folder, and checks it against
    'the shared folder list. If it is found, returns TRUE, otherwise
    'returns FALSE
If Right(strFolder, 1) <> "\" Then strFolder = strFolder + "\"
For X = 1 To frmSharedFolders.lstDirectories.ListItems.Count
    Debug.Print frmSharedFolders.lstDirectories.ListItems(X).Text; strFolder
    If UCase(frmSharedFolders.lstDirectories.ListItems(X).Text) = UCase(strFolder) Then
        IsValidSharedFolder = True: Exit Function
    End If
    If UCase(Left(strFolder, Len(frmSharedFolders.lstDirectories.ListItems(X).Text))) = UCase(frmSharedFolders.lstDirectories.ListItems(X).Text) Then
        IsValidSharedFolder = True: Exit Function
    End If
Next X
    IsValidSharedFolder = False
End Function

Sub Main()
    'Entry point for application... depending on settings
    'will display either splash screen or main window
If GetSetting("Andromeda", "Settings", "SplashScreen", "1") = "1" Then
    frmSplash.Show
    start = Timer
    Do While Timer - start < 2.5: DoEvents: Loop
    frmSplash.Hide
    frmMain.Show
Else
    frmMain.Show
End If
End Sub

Function MoveFile(oldPath As String, newPath As String, Winsock As Winsock)
    'Moves a file to new folder.
    'Sends "MOVED" to client when done, so that client may refresh file list
    On Error GoTo ErrorHandle
    Dim fSObj As FileSystemObject
    Set fSObj = CreateObject("Scripting.FileSystemObject")
    
    Call fSObj.MoveFile(oldPath, newPath) 'Move file
    
    Call Winsock.SendData("MOVED")
    
    sOutput "MOVE '" & oldPath & "' to '" & newPath & "' from IP '" & Winsock.RemoteHostIP & "'"
    Exit Function
ErrorHandle:
    Winsock.SendData "NOTMOVED"
    sOutput "Error occurred in MoveFile: " & Err.Description & " #: " & Err.Number
End Function

Function MoveFolder(oldPath As String, newPath As String, Winsock As Winsock)
    'Moves a folder and it's contents to new folder.
    'Sends "MOVED" to client when done, so that client may refresh file list
    On Error GoTo ErrorHandle
    Dim fSObj As FileSystemObject
    Set fSObj = CreateObject("Scripting.FileSystemObject")
    Dim BackSlash As Integer
    BackSlash = FindReverse(oldPath, "\")
    oldPath = Left(oldPath, BackSlash - 1)
    Call fSObj.MoveFolder(oldPath, newPath) 'Move folder
    
    Call Winsock.SendData("MOVED")
    
    sOutput "MOVE '" & oldPath & "' to '" & newPath & "' from IP '" & Winsock.RemoteHostIP & "'"
    Exit Function
ErrorHandle:
    Winsock.SendData "NOTMOVED"
    sOutput "Error occurred in MoveFolder: " & Err.Description & " #: " & Err.Number
End Function

Function FindReverse(str As String, char As String) As Integer
    'The opposite of InStr(). This function
    'will return the index of the specified character from the END
    'of the string, instead of the beginning
    ind = Len(str)
    
    Do While ind <> 1
        ch = Mid(str, ind, Len(char))
        If LCase(ch) = LCase(char) Then
            FindReverse = ind
            Exit Function
        End If
        ind = ind - 1
    Loop
    
    FindReverse = 0

End Function
Function SendDirectoryContents(Path As String, coll As Collection)
Dim objFso As New FileSystemObject
If Right(Path, 1) <> "\" Then Path = Path + "\"
    
    'This adds the files inside 'Path' (if any)
        mypath = Path
        myName = Dir(mypath)
        Do While myName <> "": DoEvents
            If myName <> "." And myName <> ".." Then
                If (GetAttr(mypath & myName) And vbDirectory) = vbDirectory Then GoTo next1
                coll.Add (mypath & myName)
                
            End If
next1:
            myName = Dir
        Loop
        
    Dim objDir1 As Folder
    Dim objDir2 As Folder
    Set objDir1 = objFso.GetFolder(Path)
    
    'This part adds all the files inside subfolders
    If objDir1.SubFolders.Count = 0 Then Exit Function
    
    For Each objDir2 In objDir1.SubFolders
        'add all the files inside the subfolder
        Call SendDirectoryContents(Path & objDir2.Name, coll)
    Next objDir2
    
    Set objDir1 = Nothing
    Set objDir2 = Nothing
    Set objFso = Nothing
     
End Function
Function ReadINI(AppName$, Keyname$, fileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   ReadINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal Keyname$, "", RetStr, Len(RetStr), fileName$))
End Function
Function ReadEncryptedINI(xAppName As String, xSubitem As String, xPathToFile As String) As String
   'Just like ReadINI, but instead, reads an encrypted entry
   'in the INI, and decrypts it before returning the value
   'Very handy :)
   ReadEncryptedINI = Decrypt(ReadINI(xAppName, xSubitem, xPathToFile))
   
End Function

Function SendProcessesToClient(Winsock As Winsock)
    'Creates a data packet that the client can
    'translate into a list of processes running on this machine
Dim xData As String

KillApp "none", frmMain.lstProcesses

For X = 0 To frmMain.lstProcesses.ListCount - 1
    xData = xData & frmMain.lstProcesses.List(X) & "|"
Next X

xData = "PROCESSES->" & xData

Winsock.SendData (xData)

sOutput "Sent processes list to IP '" & Winsock.RemoteHostIP & "' (" & Len(xData) & " Bytes)"
End Function


Function ListBox_To_String(xList As ListBox)
    'Takes a ListBox control as an argument
    'loops through the list, and concatenates the items
    'into a string separated by semicolons
On Error Resume Next

If xList.ListCount = 0 Then
    sOutput "ListBox_To_String() Returned: No items to write. Cannot continue.": Exit Function
End If
For X = 0 To xList.ListCount - 1
    Item = xList.List(X)
If X = 0 Then xData = xData & Item: GoTo aa
    xData = xData & ";" & Item
aa:
    DoEvents
Next X
    ListBox_To_String = xData
End Function

Function StartProcess(xPath As String, Winsock As Winsock)
    'Will execute a program passed in xPath
    On Error GoTo error_handle
    If GetSetting("Andromeda", "Settings", "AllowProcessToggle", "0") = "0" Then
        Winsock.SendData ("ERROR: Process toggling not allowed.")
        Exit Function
    End If
    
    'In case of malicious intent...
    If InStr(UCase(xPath), "DELTREE") <> 0 Or InStr(UCase(xPath), "FDISK") <> 0 Or InStr(UCase(xPath), "FORMAT") <> 0 Then
        'Someone thinks it would be funny to ruin the computer...
        Winsock.SendData ("ERROR: You must be stupid to attempt to run that program.")
        Exit Function
    End If
    
    'Execute the process
    Call Shell(xPath, vbNormalFocus)
    
    Winsock.SendData ("STARTED=" & xPath)
    sOutput "Started '" & xPath & "' from IP '" & Winsock.RemoteHostIP & "'"
    Exit Function
error_handle:
    Winsock.SendData ("ERROR: An error occurred while trying to spawn the process: " & xPath)
End Function

Function TerminateRunningProcess(xPath As String, Winsock As Winsock)
    'Will execute a program passed in xPath
    If GetSetting("Andromeda", "Settings", "AllowProcessToggle") = "0" Then
        Winsock.SendData ("ERROR: Process toggling not allowed.")
        Exit Function
    End If
    On Error GoTo err_handle
    KillApp xPath, frmMain.lstProcesses
    Winsock.SendData ("TERMINATED=" & xPath)
    sOutput "Terminated'" & xPath & "' from IP '" & Winsock.RemoteHostIP & "'"
    
    Exit Function
err_handle:
    Winsock.SendData ("ERROR: Process not terminated.")
    sOutput "Error in TerminateRunningProcess: xPath = " & xPath
End Function

Function WriteINI(mizainz$, Place$, Toput$, AppName$)
    r% = WritePrivateProfileString(mizainz$, Place$, Toput$, AppName$)
End Function


Function WriteEncryptedINI(xAppName As String, xSubitem As String, xOutput As String, xPathToFile As String)
    xOutput = Encrypt(xOutput)
    r% = WritePrivateProfileString(xAppName, xSubitem, xOutput, xPathToFile)
End Function

Function AppName() As String
 AppName = "ndromeda RFS "
End Function

Sub ModifyUser(xUser As String)
    'Displays a dynamically created 'frmModifyUser'
    'and initializes it's fields to the properties for the
    'specified user
    If Exists(App.Path + "\" + xUser + ".alf") = False Then MsgBox "User '" & xUser & "' does not exist.", 16, "SERVER ERROR": Exit Sub
    Dim frmModifyUser2 As New frmModifyUser
    With frmModifyUser2
    
        .txtPassword = ReadEncryptedINI("Andromeda", "PW", App.Path + "\" + xUser + ".alf")
    
        .frameUser.Caption = "User settings for: " & xUser
        
        .Caption = "ndromeda - Settings for '" & xUser & "'"
        .txtUser = xUser
        .Show
    End With
        
End Sub
Public Function Exists(fizile As String) As Boolean
    'Checks for the existence of a file or folder.
    'Returns a Boolean value (T or F)
    On Error Resume Next
    If Dir(fizile) = "" Then
        Exists = False
    Else
        Exists = True
    End If
End Function

Sub RemoveFromRegistry()
    'Deletes the key in HKEY_LOCAL_MACHINE\_
    'Software\Microsoft\Windows\CurrentVersion\Run
    '(This allows the application to be executed when
    'windows is loaded
 Dim RetVal As Long, hKey As Long, ValueName As String, _
        SubKey As String, phkResult As Long, SA As SECURITY_ATTRIBUTES, _
        Create As Long
    hKey = HKEY_LOCAL_MACHINE
    SubKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
    RetVal = RegCreateKeyEx(hKey, SubKey, _
        0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
        SA, phkResult, Create)
    ValueName = "AndromedaRFS"
    RetVal = RegDeleteValue(phkResult, ValueName)
    RegCloseKey phkResult
End Sub


Sub FileTransferAdd(xFileName As String, xFileSize As Long, xIPAddress As String, xStatus As String)
    'Adds an item to the 'File Transfer' list on the
    'main window (frmMain). When files are transferred, either
    'to or from the server, it is recorded here, and if the option
    'is enabled for logging, it is written to the file transfer log
    '(App.Path + "\FTransfer.txt")
    With frmMain.lstTransfer
        Dim pinche As ListItem
        
        Set pinche = .ListItems.Add(1, , xFileName)
        pinche.SubItems(1) = xFileSize & " bytes"
        pinche.SubItems(2) = xIPAddress
        pinche.SubItems(3) = xStatus
    
    End With
End Sub

Function Encrypt(eString As String) As String
'Takes a string as an argument,
'and encrypts it. (Doubles the memory required for the string)
Dim nextChr As String
a$ = "á0"
aa$ = "Å1"
b$ = "˚0"
bb$ = "§1"
C$ = "¸0"
cc$ = "≤1"
d$ = "ª0"
dd$ = "∂1"
e$ = "¨0"
ee$ = "Ô1"
f$ = "Ê0"
ff$ = "1∂"
g$ = "Ü0"
gg$ = "ﬁ1"
h$ = "§0"
hh$ = "É1"
i$ = "å0"
ii$ = "ˆ1"
j$ = "ô0"
jj$ = "ò1"
k$ = "x0"
KK$ = "Ê1"
l$ = "ƒ0"
ll$ = "˘1"
m$ = "∫0"
mm$ = "1%"
n$ = "≠0"
nn$ = "£1"
o$ = "¶0"
oo$ = "¯1"
p$ = "°0"
pp$ = "∆1"
q$ = "ß0"
qq$ = "≈1"
r1$ = "ã0"
rr$ = "…1"
s$ = "Õ0"
ss$ = "–1"
t$ = "«0"
tt$ = "¸1"
u$ = "—0"
uu$ = "„1"
V$ = "ä0"
vv$ = "1§"
w$ = "£0"
ww$ = "™1"
X$ = "h0"
xx$ = "√1"
Y$ = "à0"
yy$ = "1£"
z$ = "Á0"
zz$ = "®1"
qte$ = "˘0"
tld$ = "'1"
tld2$ = "∞0"
exc$ = "g1"
ats$ = "•0"
pnd$ = "ß1"
dol$ = "_0"
per$ = "Æ1"
crt$ = "ﬂ0"
amp$ = "‹1"
ast$ = "©0"
opr$ = "Ò1"
cpr$ = "0°"
dsh$ = "”1"
und$ = "◊0"
pls$ = "Í1"
eqs$ = "À0"
obc$ = "î1"
cbc$ = "∆0"
obr$ = "0∂"
cbr$ = "É0"
dsl$ = "1–"
fsl$ = "0ã"
cln$ = "˝1"
scl$ = "ì0"
fqt$ = "˜1"
apy$ = "í0"
lsn$ = "â1"
cma$ = "≥0"
grn$ = "i1"
prd$ = "π0"
qes$ = "1ò"
bsl$ = "m0"
spa$ = "w1"
zer$ = "0ª"
one$ = "ï1"
two$ = "60"
thr$ = "Ø1"
fou$ = "Â0"
fiv$ = "1p"
six$ = "∏0"
sev$ = "1Ω"
eig$ = "0£"
nin$ = "h&"
Let inptxt$ = eString
Let lenth% = Len(inptxt$)

Do While NumSpc% <= lenth%
DoEvents

Let NumSpc% = NumSpc% + 1

Let nextChr$ = Mid$(inptxt$, NumSpc%, 1)
If nextChr$ = "A" Then Let nextChr$ = aa$
If nextChr$ = "a" Then Let nextChr$ = a$
If nextChr$ = "B" Then Let nextChr$ = bb$
If nextChr$ = "b" Then Let nextChr$ = b$
If nextChr$ = "C" Then Let nextChr$ = cc$
If nextChr$ = "c" Then Let nextChr$ = C$
If nextChr$ = "D" Then Let nextChr$ = dd$
If nextChr$ = "d" Then Let nextChr$ = d$
If nextChr$ = "E" Then Let nextChr$ = ee$
If nextChr$ = "e" Then Let nextChr$ = e$
If nextChr$ = "f" Then Let nextChr$ = f$
If nextChr$ = "F" Then Let nextChr$ = ff$
If nextChr$ = "G" Then Let nextChr$ = gg$
If nextChr$ = "g" Then Let nextChr$ = g$
If nextChr$ = "H" Then Let nextChr$ = hh$
If nextChr$ = "h" Then Let nextChr$ = h$
If nextChr$ = "I" Then Let nextChr$ = ii$
If nextChr$ = "i" Then Let nextChr$ = i$
If nextChr$ = "J" Then Let nextChr$ = jj$
If nextChr$ = "j" Then Let nextChr$ = j$
If nextChr$ = "k" Then Let nextChr$ = k$
If nextChr$ = "K" Then Let nextChr$ = KK$
If nextChr$ = "L" Then Let nextChr$ = ll$
If nextChr$ = "l" Then Let nextChr$ = l$
If nextChr$ = "M" Then Let nextChr$ = mm$
If nextChr$ = "m" Then Let nextChr$ = m$
If nextChr$ = "N" Then Let nextChr$ = nn$
If nextChr$ = "n" Then Let nextChr$ = n$
If nextChr$ = "O" Then Let nextChr$ = oo$
If nextChr$ = "o" Then Let nextChr$ = o$
If nextChr$ = "P" Then Let nextChr$ = pp$
If nextChr$ = "p" Then Let nextChr$ = p$
If nextChr$ = "Q" Then Let nextChr$ = qq$
If nextChr$ = "q" Then Let nextChr$ = q$
If nextChr$ = "r" Then Let nextChr$ = r1$
If nextChr$ = "R" Then Let nextChr$ = rr$
If nextChr$ = "S" Then Let nextChr$ = ss$
If nextChr$ = "s" Then Let nextChr$ = s$
If nextChr$ = "t" Then Let nextChr$ = t$
If nextChr$ = "T" Then Let nextChr$ = tt$
If nextChr$ = "U" Then Let nextChr$ = uu$
If nextChr$ = "u" Then Let nextChr$ = u$
If nextChr$ = "V" Then Let nextChr$ = vv$
If nextChr$ = "v" Then Let nextChr$ = V$
If nextChr$ = "W" Then Let nextChr$ = ww$
If nextChr$ = "w" Then Let nextChr$ = w$
If nextChr$ = "X" Then Let nextChr$ = xx$
If nextChr$ = "x" Then Let nextChr$ = X$
If nextChr$ = "Y" Then Let nextChr$ = yy$
If nextChr$ = "y" Then Let nextChr$ = Y$
If nextChr$ = "Z" Then Let nextChr$ = zz$
If nextChr$ = "z" Then Let nextChr$ = z$
If nextChr$ = "1" Then Let nextChr$ = one$
If nextChr$ = "2" Then Let nextChr$ = two$
If nextChr$ = "3" Then Let nextChr$ = thr$
If nextChr$ = "4" Then Let nextChr$ = fou$
If nextChr$ = "5" Then Let nextChr$ = fiv$
If nextChr$ = "6" Then Let nextChr$ = six$
If nextChr$ = "7" Then Let nextChr$ = sev$
If nextChr$ = "8" Then Let nextChr$ = eig$
If nextChr$ = "9" Then Let nextChr$ = nin$
If nextChr$ = "0" Then Let nextChr$ = zer$
If nextChr$ = "~" Then Let nextChr$ = tld$
If nextChr$ = "`" Then Let nextChr$ = tld2$
If nextChr$ = "!" Then Let nextChr$ = exc$
If nextChr$ = "@" Then Let nextChr$ = ats$
If nextChr$ = "#" Then Let nextChr$ = pnd$
If nextChr$ = "$" Then Let nextChr$ = dol$
If nextChr$ = "%" Then Let nextChr$ = per$
If nextChr$ = "^" Then Let nextChr$ = crt$
If nextChr$ = "&" Then Let nextChr$ = amp$
If nextChr$ = "*" Then Let nextChr$ = ast$
If nextChr$ = "(" Then Let nextChr$ = opr$
If nextChr$ = ")" Then Let nextChr$ = cpr$
If nextChr$ = "-" Then Let nextChr$ = dsh$
If nextChr$ = "_" Then Let nextChr$ = und$
If nextChr$ = "+" Then Let nextChr$ = pls$
If nextChr$ = "=" Then Let nextChr$ = eqs$
If nextChr$ = "{" Then Let nextChr$ = obc$
If nextChr$ = "}" Then Let nextChr$ = cbc$
If nextChr$ = "[" Then Let nextChr$ = obr$
If nextChr$ = "]" Then Let nextChr$ = cbr$
If nextChr$ = "|" Then Let nextChr$ = dsl$
If nextChr$ = "\" Then Let nextChr$ = fsl$
If nextChr$ = ":" Then Let nextChr$ = cln$
If nextChr$ = ";" Then Let nextChr$ = scl$
If nextChr$ = Chr$(34) Then Let nextChr$ = qte$
If nextChr$ = "'" Then Let nextChr$ = apy$
If nextChr$ = "<" Then Let nextChr$ = lsn$
If nextChr$ = "," Then Let nextChr$ = cma$
If nextChr$ = ">" Then Let nextChr$ = grn$
If nextChr$ = "." Then Let nextChr$ = prd$
If nextChr$ = "?" Then Let nextChr$ = qes$
If nextChr$ = "/" Then Let nextChr$ = bsl$
If nextChr$ = " " Then Let nextChr$ = spa$


Let Newsent$ = Newsent$ + nextChr$
dustepp2:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Encrypt = Newsent$
End Function


Function Decrypt(dString) As String
'Takes an encrypted string (encrypted by our Encrypt() function)
'and decrypts it to normal text. See also: Read and WriteEncryptedINI()
a$ = "á0"
aa$ = "Å1"
b$ = "˚0"
bb$ = "§1"
C$ = "¸0"
cc$ = "≤1"
d$ = "ª0"
dd$ = "∂1"
e$ = "¨0"
ee$ = "Ô1"
f$ = "Ê0"
ff$ = "1∂"
g$ = "Ü0"
gg$ = "ﬁ1"
h$ = "§0"
hh$ = "É1"
i$ = "å0"
ii$ = "ˆ1"
j$ = "ô0"
jj$ = "ò1"
k$ = "x0"
KK$ = "Ê1"
l$ = "ƒ0"
ll$ = "˘1"
m$ = "∫0"
mm$ = "1%"
n$ = "≠0"
nn$ = "£1"
o$ = "¶0"
oo$ = "¯1"
p$ = "°0"
pp$ = "∆1"
q$ = "ß0"
qq$ = "≈1"
r1$ = "ã0"
rr$ = "…1"
s$ = "Õ0"
ss$ = "–1"
t$ = "«0"
tt$ = "¸1"
u$ = "—0"
uu$ = "„1"
V$ = "ä0"
vv$ = "1§"
w$ = "£0"
ww$ = "™1"
X$ = "h0"
xx$ = "√1"
Y$ = "à0"
yy$ = "1£"
z$ = "Á0"
zz$ = "®1"
qte$ = "˘0"
tld$ = "'1"
tld2$ = "∞0"
exc$ = "g1"
ats$ = "•0"
pnd$ = "ß1"
dol$ = "_0"
per$ = "Æ1"
crt$ = "ﬂ0"
amp$ = "‹1"
ast$ = "©0"
opr$ = "Ò1"
cpr$ = "0°"
dsh$ = "”1"
und$ = "◊0"
pls$ = "Í1"
eqs$ = "À0"
obc$ = "î1"
cbc$ = "∆0"
obr$ = "0∂"
cbr$ = "É0"
dsl$ = "1–"
fsl$ = "0ã"
cln$ = "˝1"
scl$ = "ì0"
fqt$ = "˜1"
apy$ = "í0"
lsn$ = "â1"
cma$ = "≥0"
grn$ = "i1"
prd$ = "π0"
qes$ = "1ò"
bsl$ = "m0"
spa$ = "w1"
zer$ = "0ª"
one$ = "ï1"
two$ = "60"
thr$ = "Ø1"
fou$ = "Â0"
fiv$ = "1p"
six$ = "∏0"
sev$ = "1Ω"
eig$ = "0£"
nin$ = "h&"

Let lenth% = Len(dString)
Let NumSpc% = 1
Do While NumSpc% <= lenth% - 1
DoEvents
Let nextChr$ = Mid$(dString, NumSpc%, 2)
Let NumSpc% = NumSpc% + 2
If nextChr$ = aa$ Then Let nextChr$ = "A"
If nextChr$ = a$ Then Let nextChr$ = "a"
If nextChr$ = bb$ Then Let nextChr$ = "B"
If nextChr$ = b$ Then Let nextChr$ = "b"
If nextChr$ = cc$ Then Let nextChr$ = "C"
If nextChr$ = C$ Then Let nextChr$ = "c"
If nextChr$ = dd$ Then Let nextChr$ = "D"
If nextChr$ = d$ Then Let nextChr$ = "d"
If nextChr$ = ee$ Then Let nextChr$ = "E"
If nextChr$ = e$ Then Let nextChr$ = "e"
If nextChr$ = f$ Then Let nextChr$ = "f"
If nextChr$ = ff$ Then Let nextChr$ = "F"
If nextChr$ = gg$ Then Let nextChr$ = "G"
If nextChr$ = g$ Then Let nextChr$ = "g"
If nextChr$ = hh$ Then Let nextChr$ = "H"
If nextChr$ = h$ Then Let nextChr$ = "h"
If nextChr$ = ii$ Then Let nextChr$ = "I"
If nextChr$ = i$ Then Let nextChr$ = "i"
If nextChr$ = j$ Then Let nextChr$ = "j"
If nextChr$ = jj$ Then Let nextChr$ = "J"
If nextChr$ = k$ Then Let nextChr$ = "k"
If nextChr$ = KK$ Then Let nextChr$ = "K"
If nextChr$ = ll$ Then Let nextChr$ = "L"
If nextChr$ = l$ Then Let nextChr$ = "l"
If nextChr$ = mm$ Then Let nextChr$ = "M"
If nextChr$ = m$ Then Let nextChr$ = "m"
If nextChr$ = nn$ Then Let nextChr$ = "N"
If nextChr$ = n$ Then Let nextChr$ = "n"
If nextChr$ = oo$ Then Let nextChr$ = "O"
If nextChr$ = o$ Then Let nextChr$ = "o"
If nextChr$ = pp$ Then Let nextChr$ = "P"
If nextChr$ = p$ Then Let nextChr$ = "p"
If nextChr$ = qq$ Then Let nextChr$ = "Q"
If nextChr$ = q$ Then Let nextChr$ = "q"
If nextChr$ = r1$ Then Let nextChr$ = "r"
If nextChr$ = rr$ Then Let nextChr$ = "R"
If nextChr$ = ss$ Then Let nextChr$ = "S"
If nextChr$ = s$ Then Let nextChr$ = "s"
If nextChr$ = t$ Then Let nextChr$ = "t"
If nextChr$ = tt$ Then Let nextChr$ = "T"
If nextChr$ = uu$ Then Let nextChr$ = "U"
If nextChr$ = u$ Then Let nextChr$ = "u"
If nextChr$ = vv$ Then Let nextChr$ = "V"
If nextChr$ = V$ Then Let nextChr$ = "v"
If nextChr$ = ww$ Then Let nextChr$ = "W"
If nextChr$ = w$ Then Let nextChr$ = "w"
If nextChr$ = xx$ Then Let nextChr$ = "X"
If nextChr$ = X$ Then Let nextChr$ = "x"
If nextChr$ = yy$ Then Let nextChr$ = "Y"
If nextChr$ = Y$ Then Let nextChr$ = "y"
If nextChr$ = zz$ Then Let nextChr$ = "Z"
If nextChr$ = z$ Then Let nextChr$ = "z"
If nextChr$ = qte$ Then Let nextChr$ = Chr$(34)
If nextChr$ = one$ Then Let nextChr$ = "1"
If nextChr$ = two$ Then Let nextChr$ = "2"
If nextChr$ = thr$ Then Let nextChr$ = "3"
If nextChr$ = fou$ Then Let nextChr$ = "4"
If nextChr$ = fiv$ Then Let nextChr$ = "5"
If nextChr$ = six$ Then Let nextChr$ = "6"
If nextChr$ = sev$ Then Let nextChr$ = "7"
If nextChr$ = eig$ Then Let nextChr$ = "8"
If nextChr$ = nin$ Then Let nextChr$ = "9"
If nextChr$ = zer$ Then Let nextChr$ = "0"
If nextChr$ = tld$ Then Let nextChr$ = "~"
If nextChr$ = tld2$ Then Let nextChr$ = "`"
If nextChr$ = exc$ Then Let nextChr$ = "!"
If nextChr$ = ats$ Then Let nextChr$ = "@"
If nextChr$ = pnd$ Then Let nextChr$ = "#"
If nextChr$ = dol$ Then Let nextChr$ = "$"
If nextChr$ = per$ Then Let nextChr$ = "%"
If nextChr$ = crt$ Then Let nextChr$ = "^"
If nextChr$ = amp$ Then Let nextChr$ = "&"
If nextChr$ = ast$ Then Let nextChr$ = "*"
If nextChr$ = opr$ Then Let nextChr$ = "("
If nextChr$ = cpr$ Then Let nextChr$ = ")"
If nextChr$ = dsh$ Then Let nextChr$ = "-"
If nextChr$ = und$ Then Let nextChr$ = "_"
If nextChr$ = pls$ Then Let nextChr$ = "+"
If nextChr$ = eqs$ Then Let nextChr$ = "="
If nextChr$ = obc$ Then Let nextChr$ = "{"
If nextChr$ = cbc$ Then Let nextChr$ = "}"
If nextChr$ = obr$ Then Let nextChr$ = "["
If nextChr$ = cbr$ Then Let nextChr$ = "]"
If nextChr$ = dsl$ Then Let nextChr$ = "|"
If nextChr$ = fsl$ Then Let nextChr$ = "\"
If nextChr$ = cln$ Then Let nextChr$ = ":"
If nextChr$ = scl$ Then Let nextChr$ = ";"
If nextChr$ = apy$ Then Let nextChr$ = "'"
If nextChr$ = lsn$ Then Let nextChr$ = "<"
If nextChr$ = cma$ Then Let nextChr$ = ","
If nextChr$ = grn$ Then Let nextChr$ = ">"
If nextChr$ = prd$ Then Let nextChr$ = "."
If nextChr$ = qes$ Then Let nextChr$ = "?"
If nextChr$ = bsl$ Then Let nextChr$ = "/"
If nextChr$ = spa$ Then Let nextChr$ = " "
Let Newsent$ = Newsent$ + nextChr$
DoEvents
Loop
Decrypt = Newsent$
End Function

Function FindPort() As Long
    'Uses a randomized seed to create an open data port
    Randomize
    FindPort = Int((10000 - 1000) * Rnd + 1000)
End Function
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
Sub TimeOut(HowLong)

'Halts program execution for a specified time (in seconds)
TheBeginning = Timer
Do While Timer - TheBeginning < HowLong
    X = DoEvents()
Loop

End Sub
Public Function KillApp(myName As String, List As ListBox) As Boolean
    'If called with "none", this function will clear frmMain.lstProcesses,
    'and then query the Windows OS for running processes, then add them
    'into the list. If called with a valid running executable's path (example: C:\MyProgram\Myprogram.exe)
    'it will terminate that process.
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    List.Clear
    
    Do While rProcessFound
        DoEvents
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        List.AddItem (szExename)
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        DoEvents
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop


    Call CloseHandle(hSnapshot)
Finish:
End Function

Function DeleteFiles(xFiles As String, IPAddress As String) As Boolean
    'Deletes files passed in xFiles... ('file1|file2|file3|')
    'Returns Boolean (True for success, False otherwise)
On Error GoTo ErrorHandle

Dim dDir As File
Dim dObj As FileSystemObject
Set dObj = CreateObject("Scripting.FileSystemObject")

If Not Mid(xFiles, Len(xFiles), 1) = "|" Then
    xFiles = xFiles & "|"
End If

For DoList = 1 To Len(xFiles)
    thechars$ = thechars$ & Mid(xFiles, DoList, 1)
    fileName = Mid(thechars$, 1, Len(thechars$) - 1)
    
    
    If Mid(xFiles, DoList, 1) = "|" Then

        If (GetAttr(fileName) And vbDirectory) = vbDirectory Then 'It's a dir, so we have to strip the .d
               
                RmDir (fileName)
                DeleteFiles = True
                sOutput "DELETE '" & fileName & "' from IP '" & IPAddress & "'"
                thechars$ = ""
        Else
        Kill fileName
        sOutput "DELETE '" & fileName & "' from IP '" & IPAddress & "'"
        thechars$ = ""
        End If
    End If
    
Next DoList

DeleteFiles = True
Exit Function
ErrorHandle:
sOutput "Error occurred in DeleteFiles(): " & Err.Description & " #: " & Err.Number
DeleteFiles = False
End Function
Sub SendFileToClient(xFileName As String, IPAddy As String, whichWinsock As Winsock)
'Opens a disk file using Binary Access Read, reads a specified block of data
'(size is governed by BufferSize (default 2048))
'after reading block of data, sends it to the remote machine
'via the Winsock control passed as the third argument (WhichWinsock)
On Error GoTo errorhandler
Dim Buffer As String
Dim BufferSize As Integer
Dim Fiz As File
Dim pinche As ListItem
Dim FizObj As Scripting.FileSystemObject
Dim fileLength As Long, SuperBuffer As Long
Dim PercentDone As Long, b As Integer

    BufferSize = 2048
            
        
         Do While whichWinsock.State <> 7: DoEvents
         If whichWinsock.State = sckError Then
         sOutput "Winsock Error:" & vbCrLf & Err.Description: Exit Sub
         End If
         Loop
         
         
         StartSending = False
         whichWinsock.SendData "FILESIZE=" & FileLen(xFileName)
         Do While StartSending <> True: DoEvents: Loop
    i = FreeFile 'Find free file
   
    
    Set FizObj = CreateObject("Scripting.FileSystemObject")
    Set Fiz = FizObj.GetFile(xFileName)
    
    Set pinche = frmMain.lstTransfer.ListItems.Add(1, , Fiz.ParentFolder + "\" + Fiz.Name)
        pinche.SubItems(1) = Fiz.Size & " bytes"
        pinche.SubItems(2) = IPAddy
        
    Open xFileName For Binary Access Read As #i
    
        fileLength = LOF(i)
       
        Do While Not EOF(i): DoEvents
        
            If fileLength - Loc(i) < BufferSize Then
                Let BufferSize = fileLength - Loc(i)
                If BufferSize = 0 Then GoTo done
            End If
            
            Buffer = Space(BufferSize)
       
        If Loc(i) = 0 Then GoTo skipPercent 'Don't want division by zero
        
        PercentDone = Loc(i) / fileLength * 100
        
        If b < 30 Then: b = b + 1: GoTo skipPercent
        pinche.SubItems(3) = PercentDone & "%"
        b = 0
skipPercent:
    
        Get #i, , Buffer
        
     
        whichWinsock.SendData Buffer

        SuperBuffer = SuperBuffer + Len(Buffer)
        
        Loop
done:
    Close #i
        StartSending = False
        pinche.SubItems(3) = "Complete."
        sOutput "SENT-> " & xFileName & " (" & SuperBuffer & " bytes) to [" & IPAddy & "]"
        If GetSetting("Andromeda", "Settings", "WriteTransferLog") = "1" Then
            WriteLog App.Path + "\FTransfer.txt", "Sent '" & xFileName & "' (" & SuperBuffer & " bytes) to IP '" & IPAddy & "' Time/Date=" & Format(Now, "HH:MM:SS AM/PM - MM/DD/YYYY")
        End If
        Exit Sub
errorhandler:
Call sOutput("Error in SendFileToClient: " & Err.Description & "Number: " & Err.Number)

End Sub

Sub LoadExistingUserInformation()

    With frmManageUsers
        Path$ = App.Path + "\*.alf"
        firsthtml = Dir(Path$)
        If firsthtml = "" Then
            MsgBox "No User files could be found to search through.", 16, "Files Not Found": btnCancel.Enabled = False:  Exit Sub
        End If
        
        .lstbuffer.Clear

    Do While firsthtml <> "": DoEvents
        .lstbuffer.AddItem (firsthtml)
        firsthtml = Dir
    Loop
        .ListView1.ListItems.Clear
        
    For X = 0 To .lstbuffer.ListCount - 1
        username = ReadEncryptedINI("Andromeda", "UserName", App.Path + "\" + .lstbuffer.List(X))
        Pw = ReadEncryptedINI("Andromeda", "PW", App.Path + "\" + .lstbuffer.List(X))
        lastlogin = ReadEncryptedINI("Andromeda", "LastLogin", App.Path + "\" + .lstbuffer.List(X))
Dim itm As ListItem

        Pw = "''" & Pw & "''"
        
        If lastlogin = "" Then lastlogin = "Never."
        
       
If username = "" Then GoTo izend
Set itm = .ListView1.ListItems.Add(, , username)
    itm.SubItems(1) = Pw
    itm.SubItems(2) = lastlogin
    
    DoEvents
izend:
    Next X
    
    End With
End Sub
Sub LoadSharedDirectories()
    'Opens the Shared Directories config file (\SD.DLL)
    'and adds the shared folders to frmSharedFolders.lstDirectories
    Dim Jig As ListItem
    With frmSharedFolders
    i = FreeFile
    If Exists(App.Path + "\SD.DLL") = False Then
        MsgBox "No shared directory information could be located. You need to add shared directories.", 16, "Error: No Shared Directory Information": Exit Sub
    End If
        Open App.Path + "\SD.DLL" For Input As #i
            Do While Not EOF(i): DoEvents
            
            Line Input #i, shmoo$
            
            Set Jig = .lstDirectories.ListItems.Add(, , shmoo$, , 1)
            
            Loop
        
        Close #i
    End With
End Sub

 Sub WriteRegistry(hKey As Long, SubKey As String, _
    ValueName As String, vNewValue As String)
    Dim phkResult As Long, RetVal As Long
    'Writes a key in the registry under the HKEY passed in first argument
    RetVal = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, phkResult)
    RetVal = RegSetValueEx(phkResult, ValueName, 0, REG_SZ, vNewValue, _
          CLng(Len(vNewValue) + 1))
    
         
    'Close the keys
    RegCloseKey hKey
    RegCloseKey phkResult

End Sub

Sub SaveSharedDirectories()
    'Saves the list of shared directories for the server
    i = FreeFile
    Open App.Path + "\SD.DLL" For Output As #i
        For X = 1 To frmSharedFolders.lstDirectories.ListItems.Count
        Print #i, frmSharedFolders.lstDirectories.ListItems(X).Text
        DoEvents
        Next X
    Close #i
End Sub
Function DirectoryToString(xPath As String) As String
    'Takes the path to a folder as an argument,
    'and returns a string that the client can translate
    'into a list of files and subdirectories
On Error GoTo ErrorHandle
    Dim FirstFile, xPath2 As String
    Dim fizile As File
    Dim ReturnValue As String
    Dim FizileObject As Scripting.FileSystemObject
    Set FizileObject = CreateObject("Scripting.FileSystemObject")
    

'----------------------------------- Directories
    mypath = xPath
    If mypath = "C:\" Then myName = Dir(mypath, vbDirectory): GoTo skip1
    myName = Dir(xPath, vbDirectory)
skip1:
    Do While myName <> ""
  
   If InStr(mypath, myName) = 0 Then
   If myName <> "." And myName <> ".." Then

      If (GetAttr(mypath & myName) And vbDirectory) = vbDirectory Then
         ReturnValue = ReturnValue & myName & ".d|"
      End If
   End If
   End If
skip:
   myName = Dir
Loop
    
'------------------------------------ Files
    xPath2 = xPath & "*.*"
    If xPath = "C:\" Then xPath2 = xPath
    FirstFile = Dir$(xPath2)

Do While FirstFile <> "": DoEvents
    
    Set fizile = FizileObject.GetFile(xPath + FirstFile)
    ReturnValue = ReturnValue & fizile.Name & ":" & fizile.Size & "|"
    FirstFile = Dir
Loop
    DirectoryToString = ReturnValue
Exit Function
ErrorHandle:
sOutput "Error occurred in DirectoryToString(): " & Err.Description & " #: " & Err.Number
End Function

Sub sOutput(xOutput As String)
    'Displays output in the Server Output list on frmMain
    'Called to alert user to activity
    Dim Dta As String, LITM As ListItem
    Set LITM = frmMain.lstOutput.ListItems.Add(1, , xOutput)
        LITM.SubItems(1) = Format(Now, "HH:MM:SS AM/PM - MM/DD/YYYY")
    Call WriteLog(App.Path + "\Output.txt", xOutput & " : " & Format(Now, "HH:MM:SS AM/PM - MM/DD/YYYY"))
End Sub


Sub WriteLog(strPath As String, strLine As String)
    'Writes data (strLine) to a specified file (strPath)
    '*Only appends to the end of the file, does not erase any
    'existing data in the file*
    If Exists(strPath) = False Then
        i = FreeFile
        Open strPath For Output As #i
        Close #i
    End If

    Dim fileNum As Integer
    
    fileNum = FreeFile
    
    Open strPath For Append As #fileNum
    
        Print #fileNum, strLine$ '<- for some reason, if you dont include the '$' char, it adds quotes to the line written... weird
    
    Close #fileNum
End Sub


