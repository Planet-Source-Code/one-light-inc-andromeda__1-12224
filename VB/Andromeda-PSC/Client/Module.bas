Attribute VB_Name = "Module"
Public DirPath As String
Public CurrentServer As ServerInfo

Public strBuffer As String

Public CurrentFile As String
Public fComplete As Boolean

Public FileNum As Integer
Public FileSize As Long

Public StartSending As Boolean
Public Sending As Boolean
Public FileDone As Boolean

Public WaitForMove As Boolean
Public WaitForFolder As Boolean

Public CancelDownload As Boolean
Public CancelUpload As Boolean

Public WaitForServerRecieve As Boolean

Public WaitingForContents As Boolean

Public objFso As New FileSystemObject

Public Type ServerInfo
    ServerLabel As String
    ServerIP As String
    Login As String
    Password As String
    InitDir As String
End Type

''''''''' Options ''''''''''''''''''''''''''
Public AutoReconnect As Integer
Public AutoshowFileTransfer As Integer
Public ShowSplash As Integer
Public ShowServerList As Integer
''''''''''''''''''''''''''''''''''''''''''''

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Const GW_CHILD = 5

Public Const WM_GETTEXTLENGTH = &HE

Public Const GW_HWNDNEXT = 2





Function DLL() As String
    Path$ = App.Path + "\Andromeda.cfg"
    
    If Exists(Path$) = False Then
        i = FreeFile
        Open Path$ For Output As #i
            Print #i, "[Andromeda]"
            Print #i, "AutoReconnect=1"
            Print #i, "AutoshowFileTransfer=1"
            Print #i, "ShowSplash=1"
            Print #i, "ShowServerList=1"
        Close #i
    End If
    
    DLL = Path$
End Function

Function Exists(strPath As String) As Boolean
    If Dir(strPath) = "" Then
        Exists = False
    Else
        Exists = True
    End If
End Function


Function FindChildByClass(parentw, childhand)
c% = GetWindow(parentw, GW_CHILD)
While c%
    DoEvents
    a% = SendMessage(c%, WM_GETTEXTLENGTH, 0, 0)
    b$ = String$(255, 0)
    g% = GetClassName(c%, b$, 255)
    b$ = Left(b$, g%)
    If UCase$(b$) = UCase$(childhand) Then
        FindChildByClass = c%
        Exit Function
    End If
    c% = GetWindow(c%, GW_HWNDNEXT)
Wend
End Function

Function AppName() As String
    AppName = "Andromeda"
End Function


Sub CenterFormMDI(MDI_Form As Form, Form As Form)
    Form.Left = MDI_Form.Width / 2 - Form.Width / 2
    Form.Top = MDI_Form.Height / 2 - Form.Height / 2 - 615
End Sub



Sub CenterForm(Form As Form)
    Form.Left = Screen.Width / 2 - Form.Width / 2
    Form.Top = Screen.Height / 2 - Form.Height / 2 - 615
End Sub

Function GetRunningProcesses() As String
    frmMain.Winsock.SendData ("GETPROCESSES")
    
End Function
Function InitializeProcessList(strData As String)
    'Will actually add the process list from the server
    'into the list box on frmProcesses
    For x = 1 To Len(strData)
        If Not Mid(strData, x, 1) = "|" Then
            strtemp = strtemp + Mid(strData, x, 1)
        Else
            frmProcesses.lstServersProcesses.AddItem strtemp
            strtemp = ""
        End If
    Next x
End Function

Function IsValidFileName(strName As String) As Boolean
    If InStr(strName, "\") Or _
    InStr(strName, "/") Or _
    InStr(strName, "?") Or _
    InStr(strName, ":") Or _
    InStr(strName, "*") Or _
    InStr(strName, "?") Or _
    InStr(strName, Chr(34)) Or _
    InStr(strName, "<") Or _
    InStr(strName, ">") Or _
    InStr(strName, "|") Then
        IsValidFileName = False
    Else
        IsValidFileName = True
    End If
End Function



Sub Main()
    
    'Load options into variables
    AutoReconnect = CInt(ReadINI("Andromeda", "AutoReconnect", DLL()))
    AutoshowFileTransfer = CInt(ReadINI("Andromeda", "AutoshowFileTransfer", DLL()))
    ShowSplash = CInt(ReadINI("Andromeda", "ShowSplash", DLL()))
    ShowServerList = CInt(ReadINI("Andromeda", "ShowServerList", DLL()))
    
    
    If ShowSplash = 1 Then
        frmSplash.Show
        Timeout 3
        Unload frmSplash
    End If
    
    'Show MDI form
    frmMain.Show
End Sub


Sub Timeout(interval)
    current = Timer
    Do While Timer - current < Val(interval)
    DoEvents
    Loop
End Sub
Function MDI() As Integer
    'returns the handle of frmMain's MDI area
    MDI = FindChildByClass(frmMain.hwnd, "MDIClient")
End Function



Sub ServerError(strDescription As String)
    Dim errWnd As New frmError
    errWnd.Show
    errWnd.txtDescription = strDescription
    
    Do While errWnd.Visible = True: DoEvents: Loop
End Sub

Sub ShowSharedFolders(theitems As String)
    On Error Resume Next
    
    If Not Mid(theitems, Len(theitems), 1) = "|" Then
    theitems = theitems & "|"
    End If
    
    For DoList = 1 To Len(theitems)
    thechars$ = thechars$ & Mid(theitems, DoList, 1)
    
    If Mid(theitems, DoList, 1) = "|" Then
    Call frmFileView.ServerDrives.ComboItems.Add(, , Mid(thechars$, 1, Len(thechars$) - 1), "folder")
    
    thechars$ = ""
    If Mid(theitems, DoList + 1, 1) = " " Then
    DoList = DoList + 1
    End If
    End If
    Next DoList
    
    If CurrentServer.InitDir <> "" Then
        For x = 1 To frmFileView.ServerDrives.ComboItems.Count
            If LCase(frmFileView.ServerDrives.ComboItems(x)) = LCase(CurrentServer.InitDir) Then
                frmFileView.ServerDrives.ComboItems(x).Selected = True
                GoTo skipit
            End If
        Next x
        frmFileView.ServerDrives.ComboItems.Add(, , CurrentServer.InitDir, "folder").Selected = True
skipit:
    Else
        frmFileView.ServerDrives.ComboItems(1).Selected = True
    End If
    
    Path = frmFileView.ServerDrives.SelectedItem.Text
    If Path = "" Then Exit Sub
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    
    Call frmMain.Winsock.SendData("DIR " + Path)
    ServerPath = Path
    frmFileView.ServerFileList.MousePointer = 13
    frmFileView.ServerFileList.ListItems.Clear
    frmFileView.ServerFileList.SetFocus
End Sub

Sub StringToDir(FileList As ListView, theitems As String)
FileList.ListItems.Clear
Dim TMP As ListItem
If theitems = "" Then Exit Sub
If Not Mid(theitems, Len(theitems), 1) = "|" Then
theitems = theitems & "|"
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "|" Then

txt = Mid(thechars$, 1, Len(thechars$) - 1)
cln = InStr(txt, ":")
If cln <> 0 Then
    filename = Left(txt, cln - 1)
    fsize = Right(txt, Len(txt) - cln)
Else
    filename = txt
    fsize = ""
End If
Dim ext As String


Dim f As Scripting.FileSystemObject
Set f = CreateObject("Scripting.FileSystemObject")
ext = f.GetExtensionName(filename)

    If ext = "d" Then
        rep = "folder"
        rep2 = " Directory"
        filename = f.GetBaseName(filename)
    Else
        rep = GetImage(ext)
        rep2 = GetType(ext)
    End If


Set TMP = FileList.ListItems.Add(, , filename, rep, rep)
    If rep <> "folder" And fsize <> "" Then TMP.SubItems(1) = FormatFileSize(CLng(fsize))
    TMP.SubItems(2) = rep2


next1:
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Function FormatTime(Seconds As String) As String
Dim Secs, Mins, Hours, Days
Dim TotalMins, TotalHours, TotalSecs, TempSecs

    TotalSecs = Int(Seconds)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)


    If TotalHours > 23 Then
        Hours = (TotalHours - 23)
    Else
        Hours = TotalHours
    End If


    If TotalMins > 59 Then
        Mins = (TotalMins - (Hours * 60))
    Else
        Mins = TotalMins
    End If

    If Days <> 0 Then FormatTime = Days & " days " & Hours & " hours " & Mins & " minutes " & TotalSecs & " seconds": Exit Function
    If Hours <> 0 Then FormatTime = Hours & " hours " & Mins & " minutes " & TotalSecs & " seconds": Exit Function
    If Mins <> 0 Then FormatTime = Mins & " minutes " & TotalSecs & " seconds": Exit Function
    If Secs <> 0 Then FormatTime = TotalSecs & " seconds": Exit Function
End Function

Function FormatFileSize(bytes As Long) As String
    If bytes < 1024 Then
        m = " bytes"
        FormatFileSize = bytes & m
        Exit Function
    End If
    
    If bytes >= 1024 And bytes < 1024000 Then
        m = " Kb"
        FormatFileSize = FormatNumber(bytes / 1024, 2, , , vbTrue) & m
        Exit Function
    End If
    
    If bytes >= 1024000 And bytes < 1024000000 Then
        m = " Mb"
        FormatFileSize = FormatNumber(bytes / 1024000, 2, , , vbTrue) & m
        Exit Function
    End If
    
    If bytes >= 1024000000 Then
        m = " Gb"
        FormatFileSize = FormatNumber(bytes / 1024000000, 2, , , vbTrue) & m
        Exit Function
    End If
    
End Function




