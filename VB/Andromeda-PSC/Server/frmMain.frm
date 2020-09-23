VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   5925
   ClientLeft      =   7710
   ClientTop       =   6165
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstProcesses 
      Height          =   840
      Left            =   5160
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSWinsockLib.Winsock UP 
      Index           =   0
      Left            =   5880
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock UDP 
      Index           =   0
      Left            =   6360
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   7215
   End
   Begin MSComctlLib.ListView lstTransfer 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   4975
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstOutput 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Details"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time/Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ListBox lstSharedFolders 
      Height          =   840
      Left            =   5520
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer TimerUptime 
      Interval        =   1000
      Left            =   6360
      Top             =   5160
   End
   Begin VB.Frame FrameStatus 
      Caption         =   "Server Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   3255
      Begin VB.Label txtElapsed 
         Caption         =   "00:00:00"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblUptimeasdf 
         Caption         =   "Uptime:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblServerRunning 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Server Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblConnections 
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Current Users:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame frameTop 
      Height          =   25
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5880
      Top             =   5160
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   6840
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4800
      Picture         =   "frmMain.frx":014A
      Top             =   4680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTransfer 
      Caption         =   "File Transfer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblOutput 
      Caption         =   "Server Output:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu File_MNU 
      Caption         =   "&File"
      Begin VB.Menu StartServer_MNU 
         Caption         =   "Start Server"
         Shortcut        =   {F1}
      End
      Begin VB.Menu TerminateServer_MNU 
         Caption         =   "Terminate Server"
         Shortcut        =   {F2}
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu TerminateandExit_MNU 
         Caption         =   "&Exit Andromeda RFS"
      End
   End
   Begin VB.Menu Users_MENU 
      Caption         =   "&Users"
      Begin VB.Menu MUsers_MENU 
         Caption         =   "&Manage Users"
         Shortcut        =   ^U
      End
      Begin VB.Menu CreateNewUser_MENU 
         Caption         =   "&Create New User..."
         Shortcut        =   ^N
      End
      Begin VB.Menu Logs_MENU 
         Caption         =   "&Server Logs"
         Begin VB.Menu ServerOutputHistory_MENU 
            Caption         =   "View Server Output History..."
         End
         Begin VB.Menu LoginHistory_MENU 
            Caption         =   "&View User Login History"
         End
         Begin VB.Menu TransferLog_MENU 
            Caption         =   "View File Transfer Log"
         End
      End
   End
   Begin VB.Menu Tools_MENU 
      Caption         =   "&Tools"
      Begin VB.Menu ManageDirs_MENU 
         Caption         =   "Manage &Shared Directories"
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu Config_MENU 
         Caption         =   "Server &Configuration"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu Help_MENU 
      Caption         =   "&Help"
      Begin VB.Menu About_MENU 
         Caption         =   "About Andromeda RFS"
      End
      Begin VB.Menu WebSite_MENU 
         Caption         =   "Visit Web Site"
      End
   End
   Begin VB.Menu MNU 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu ccommands_MENU 
         Caption         =   "&Shared Directories..."
      End
      Begin VB.Menu ManageUsers_MENU 
         Caption         =   "&Manage Users..."
      End
      Begin VB.Menu CNewUser_MENU 
         Caption         =   "&Create New User..."
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu ServerLogs_MENU 
         Caption         =   "Server &Logs"
         Begin VB.Menu FileTransferHistory_MENU 
            Caption         =   "View File Transfer History..."
         End
         Begin VB.Menu VLoginHistory_MENU 
            Caption         =   "View &Login History..."
         End
         Begin VB.Menu VOutputLog_MENU 
            Caption         =   "View Server &Output Log..."
         End
      End
      Begin VB.Menu State_MENU 
         Caption         =   "Server State"
         Begin VB.Menu Enable_MENU 
            Caption         =   "Enable Server"
         End
         Begin VB.Menu Disable_MENU 
            Caption         =   "Disable Server"
         End
      End
      Begin VB.Menu sep10 
         Caption         =   "-"
      End
      Begin VB.Menu TProcess_MENU 
         Caption         =   "&Exit Andromeda RFS"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intMax As Integer
Public intMax2 As Integer
Public ttlLogins As Integer






Private Sub About_MENU_Click()
frmAbout.Show , Me
End Sub

Private Sub ccommands_MENU_Click()
frmSharedFolders.Show
End Sub

Private Sub CNewUser_MENU_Click()
frmCreateUser.Show
End Sub

Private Sub Config_MENU_Click()
frmOptions.Show vbModal, Me
End Sub

Private Sub CreateNewUser_MENU_Click()
frmCreateUser.Show vbModal, Me
End Sub

Private Sub Disable_MENU_Click()
If sEnabled = False Then MsgBox "Server already stopped.", 16, "Error": Exit Sub
Call EnableServer(False)
End Sub

Private Sub Enable_MENU_Click()
If sEnabled = True Then MsgBox "Server already started.", 16, "Error": Exit Sub
Call EnableServer(True)
End Sub

Private Sub FileTransferHistory_MENU_Click()
DisplayLogFile "FileTransfer"
End Sub

Private Sub Form_Load()

If GetSetting("Andromeda", "Settings", "FirstRun") = "" Then
    'This is the first time the program has been run
    'Save default settings to registry, and inform user
    'that they must specify directories and users before
    'using the server
    SaveSetting "Andromeda", "Settings", "SplashScreen", "1"
    SaveSetting "Andromeda", "Settings", "MinimizeToTray", "0"
    SaveSetting "Andromeda", "Settings", "WriteLog", "1"
    SaveSetting "Andromeda", "Settings", "WriteTransferLog", "1"
    SaveSetting "Andromeda", "Settings", "AllowDelete", "0"
    SaveSetting "Andromeda", "Settings", "AllowMove", "1"
    SaveSetting "Andromeda", "Settings", "AllowRename", "1"
    SaveSetting "Andromeda", "Settings", "StartWithWindows", "0"
    SaveSetting "Andromeda", "Settings", "FirstRun", "1"
    SaveSetting "Andromeda", "Settings", "AllowProcessToggle", "1"
    MsgBox "Welcome to Andromeda RFS v1.0!" & vbCrLf & vbCrLf & "Since this is the first time you have started the server, you must specify shared directories by clicking on the Tools menu, and choosing 'Manage Shared Directories'. This will display the shared folder window. Add the folders you wish to share with Andromeda clients. You will also need to create user accounts for anyone that wishes to connect to your computer. You can do this by clicking on the Users menu, and choosing 'Create New User'." & vbCrLf & vbCrLf & "Thanks for trying Andromeda RFS!" & vbCrLf & "Ryan and Andrew Lederman", 64, "IMPORTANT NOTE"
End If

    If FileLen(App.Path + "\sd.dll") = 0 Then
        RetVal = MsgBox("You do not have any shared directories! Andromeda clients will not be able to access your files. It is suggested that you add some shared directories now. Would you like to add some shared directories?", 36, "No Shared Directories")
        Select Case RetVal
            Case vbYes
                frmSharedFolders.Show , Me
        End Select
    End If
    
    'Initialize the command accepting winsock on port 6969
    Server(0).LocalPort = 6969
    Server(0).Listen
    
    UP(0).LocalPort = 6971
    UP(0).Listen
   
    intMax2 = 0
    Me.Caption = AppName() & "(Enabled)"
    lblServerRunning.Caption = Server(0).LocalIP
    sEnabled = True
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .szTip = "Andromeda RFS 1.0" & vbNullChar
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Image1.Picture
    End With
    
    Shell_NotifyIcon NIM_ADD, nid
    
    KillApp "none", lstProcesses 'Initialize list of processes
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'On Error Resume Next
    'This procedure receives the callbacks from the System Tray icon.
    '
    Dim Result As Long
    Dim msg As Long
    'The value of X will vary depending upon the scalemode setting


    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If

    Select Case msg
        Case 517  '517 display popup menu

            Me.PopupMenu MNU
      
        Case 514
        Result = SetForegroundWindow(Me.hwnd)
        Me.WindowState = vbNormal
        Me.Show
    End Select

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then
    If GetSetting("Andromeda", "Settings", "MinimizeToTray", "0") = "1" Then
        Me.Hide
    End If
    End If
    
    lstOutput.Width = Me.Width - 320
    lstTransfer.Width = Me.Width - 320
    frameTop.Width = Me.Width + 300
      FrameStatus.Top = Me.Height - FrameStatus.Height - 760
    Frame1.Width = Me.Width - 320
    lstOutput.Height = Me.Height / 2 - 1350
    Frame1.Top = lstOutput.Top + lstOutput.Height + 120
    lblTransfer.Top = Frame1.Top + 80
  
    lstTransfer.Top = lblTransfer.Top + lblTransfer.Height + 20
    lstTransfer.Height = FrameStatus.Top - lstTransfer.Top - 20
    lstConnections.ColumnHeaders(4).Width = lstConnections.Width - lstConnections.ColumnHeaders(1).Width - lstConnections.ColumnHeaders(2).Width - lstConnections.ColumnHeaders(3).Width - 95
    lstOutput.ColumnHeaders(1).Width = lstOutput.Width - 2000
    lstOutput.ColumnHeaders(2).Width = lstOutput.Width - lstOutput.ColumnHeaders(1).Width - 90
    
End Sub


Private Sub lstConnections_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstConnections.SortKey = ColumnHeader.Index - 1
    lstConnections.Sorted = True
End Sub


Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
End Sub

Private Sub LoginHistory_MENU_Click()
Call DisplayLogFile("Login")
End Sub

Private Sub ManageDirs_MENU_Click()
frmSharedFolders.Show vbModal, Me
End Sub

Private Sub ManageUsers_MENU_Click()
frmManageUsers.Show , Me
End Sub

Private Sub MUsers_MENU_Click()
frmManageUsers.Show , Me
End Sub

Private Sub Server_Close(Index As Integer)
    If Index = 0 Then GoTo skip
    If ttlLogins = 0 Then GoTo skip
    ttlLogins = ttlLogins - 1
    Server(Index).Close
skip:
    
End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    'If Server(Index).State <> sckClosed Then Server(Index).Close
    'Server(Index).Accept (requestID)
    If Index = 0 Then
        dcount = Server.Count + 1
        Load Server(dcount)
        Server(dcount).Accept requestID
        Server(dcount).SendData ("LOGIN")
    End If
    

   
End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim data As String, fileName As String, renameTo As String
    Call Server(Index).GetData(data, vbString, bytesTotal)

    
    If Left(data, 6) = "LOGIN=" Then
        'Client sent login information
        colon = InStr(data, ":")
        If colon = 0 Then
            frmMain.Server(Index).SendData ("INVALID_LOGIN")
        Exit Sub
        End If
        temp = Right(data, Len(data) - ln - 6)
        colon = InStr(temp, ":")
        
        'Intialize username and password variables
        Login = Left(temp, colon - 1)
        Password = Right(temp, Len(temp) - colon)
    
    'Check for existence of %Login%.alf in the application directory
    If Exists(App.Path + "\" + Login + ".alf") = True Then
     
        'Read encrypted password from user's configuration file
        If Password = ReadEncryptedINI("Andromeda", "PW", App.Path + "\" + Login + ".alf") Then
            'Login accepted, send welcome
            Call Server(Index).SendData("WELCOME")
            
            'Increment active socket connections count
            ttlLogins = ttlLogins + 1
            
            'Write the login time to the users configuration file
            'to use for last login time/date reference
            Call WriteEncryptedINI("Andromeda", "LastLogin", Format(Now, "MM/DD/YY - HH:MM:SS AM/PM"), App.Path + "\" + Login + ".alf")
            
            'If logging is enabled in options, write the login event
            'to log
            If GetSetting("Andromeda", "Settings", "WriteLog") = "1" Then
                WriteLog App.Path + "\Log.txt", "User '" & Login & "' logged in from IP '" & Server(Index).RemoteHostIP & "' Time/Date: '" & Format(Now, "mm/dd/yy - HH:MM:SS AM/PM")
            End If
                
            sOutput "User '" & Login & "' logged in from IP '" & Server(Index).RemoteHostIP & "'"
       Else
            'Password did not match, inform client
            Call Server(Index).SendData("INVALID_PASSWORD")
            sOutput "Invalid Password for '" & Login & "' : (" & Password & ") from IP '" & Server(Index).RemoteHostIP & "'"
            Login = "": Password = ""
            Exit Sub
        End If
    Else
        'Login name was not found
        Call Server(Index).SendData("INVALID_LOGIN")
        Call Server(Index).SendData(InvalidMessage)
        sOutput "Invalid Login received: '" & Login & "' from IP '" & Server(Index).RemoteHostIP & "'"
        Login = "": Password = ""
        Exit Sub
    End If
        Login = "": Password = ""
    Exit Sub
    End If
    
   
    
    If data = "BEGIN" Then
        'Client has informed server to start sending the data
        StartSending = True
    Exit Sub
    End If
    
    If data = "FILEUPPORT" Then
        'Client has asked for a port to upload a file
        'Find an open data port, inform the client that it should
        'send the data to that port, then create a new socket and
        'attach it to that port
        Dim Piz As Long
            Piz = FindPort
            intMax2 = intMax2 + 1
            Load frmMain.UP(intMax2)
            frmMain.UP(intMax2).LocalPort = Piz
            frmMain.UP(intMax2).Listen
             Do While frmMain.UP(intMax2).State <> 2: DoEvents
            Loop
            frmMain.Server(Index).SendData ("FILEUPPORT=" & Piz)
    Exit Sub
    End If
    
    If data = "FILEPORT" Then
        'Request for open file transfer port
        Dim Piz2 As Long
        Piz2 = FindPort
        frmMain.UDP(Index).LocalPort = Piz2
        frmMain.UDP(Index).Listen
        Do While frmMain.UDP(Index).State <> 2: DoEvents
        Loop
        frmMain.Server(Index).SendData ("FILEPORT=" & Piz2)
    End If
    
     If Left(data, 17) = "GET_DIR_CONTENTS=" Then
        'Request for folder contents to download
        'all files and subdirectories inside a folder
        Dim Col1 As New Collection
        Equals2 = InStr(data, "=")
        whichFolder2$ = Right(data, Len(data) - Equals2)
        Call SendDirectoryContents(whichFolder2$, Col1)
        If Col1.Count = 0 Then
            'Some error must have occurred
            Server(Index).SendData ("ERROR:An error occurred while trying to get the contents of that directory.")
            Exit Sub
        End If
        
        For X2 = 1 To Col1.Count
            temp = temp + Col1(X2) + "|"
        Next X2
       
        Server(Index).SendData ("DIR_CONTENTS->" & temp)
    
        sOutput "Contents of '" & whichFolder2$ & "' sent to IP '" & Server(Index).RemoteHostIP & "'"
        
        Exit Sub
      End If
      
    If Left(data, 12) = "GETPROCESSES" Then
        'Request from client for list of processes
        Call SendProcessesToClient(Server(Index))
        Exit Sub
    End If
    
    If Left(data, 14) = "BEGIN_DOWNLOAD" Then
        'Create new file sending winsock
        newudp = UDP.Count
        Load UDP(newudp)
        UDP(newudp).Close
        Call UDP(newudp).Connect(Server(Index).RemoteHostIP, 109)
        Exit Sub
    End If
    
   
    
        
        
      If Left(data, 7) = "DELETE=" Then
        'Request from client to delete a file
        'check if deletion is allowed in options
        If GetSetting("Andromeda", "Settings", "AllowDelete") = "0" Then
            Server(Index).SendData ("ERROR:Andromeda RFS v1.0" & vbCrLf & "Error: Deletion not allowed!")
            Exit Sub
        End If
            
        'Call DeleteFiles() to delete the potentially large list of files
        Dim dFiles As String, dSuccess As Boolean
        dFiles = Right(data, Len(data) - InStr(data, "DELETE=") - 6)
        
        dSuccess = DeleteFiles(dFiles, Server(Index).RemoteHostIP)
      
        If dSuccess Then
            'Inform client that the file(s) were deleted
            Server(Index).SendData ("DELETED")
        End If
        Exit Sub
      End If
      
      If Left(data, 10) = "NEWFOLDER=" Then
        'Request for creation of new directory
        'Call CreateFolder()
        Dim xFolder As String, EqualIndex As Integer
        EqualIndex = InStr(data, "=")
        xFolder = Right(data, Len(data) - EqualIndex)
        Call CreateFolder(xFolder, Server(Index))
        Exit Sub
      End If
      
      If Left(data, 5) = "MOVE=" Then
        'Request to move a directory or file(s)
      If GetSetting("Andromeda", "Settings", "AllowMove") = "0" Then
            Server(Index).SendData ("ERROR:Andromeda RFS v1.0" & vbCrLf & "Error: Moving not allowed!")
            Exit Sub
      End If
        
        Dim Equals As Integer, Pipe As Integer, ExistingFile As String, NewLocation As String
        Equals = InStr(data, "=")
        data = Right(data, Len(data) - Equals)
        Pipe = InStr(data, "|")
        ExistingFile = Left(data, Pipe - 1)
        NewLocation = Right(data, Len(data) - Pipe)
        
        On Error GoTo errorInMove
        If (GetAttr(ExistingFile) And vbDirectory) = vbDirectory Then
            'Argument passed from client is a folder, move it
            Call MoveFolder(ExistingFile, NewLocation, Server(Index))
            Server(Index).SendData ("MOVED")
            Exit Sub
        Else
            'Argument passed is a file, move it
            Call MoveFile(ExistingFile, NewLocation, Server(Index))
            Server(Index).SendData ("MOVED")
            Exit Sub
        End If
        Exit Sub
errorInMove:
        Server(Index).SendData ("NOTMOVED")
        Server(Index).SendData ("ERROR:Andromeda RFS 1.0" & vbCrLf & "Error occurred in move")
        Exit Sub
      End If
        
      If Left(data, 7) = "RENAME=" Then 'Rename file

      If GetSetting("Andromeda", "Settings", "AllowRename") = "0" Then
            Server(Index).SendData ("ERROR:Andromeda RFS v1.0" & vbCrLf & "Error: Rename not allowed!")
            Exit Sub
      End If
      
            Rename = InStr(data, "RENAME=")
            nextStr = Right(data, Len(data) - Rename - 6)
            Equals = InStr(nextStr, "|")
            fileName = Left(nextStr, Equals - 1)
            renameTo = Right(nextStr, Len(nextStr) - Equals)
            torf = RenameFile(fileName, renameTo)
            
            If (GetAttr(ExistingFile) And vbDirectory) = vbDirectory Then
                'Argument passed from client is a folder, move it
                Call RenameFolder(fileName, renameTo)
             Else
                 'Argument passed is a file, move it
                Call RenameFile(fileName, renameTo)
            End If
            
            sOutput "RENAME '" & fileName & "' to '" & renameTo & "' from IP '" & Server(Index).RemoteHostIP & "'"
            
        If torf = True Then
            Server(Index).SendData ("RENAMED")
        End If
        Exit Sub
    End If
    
    If Left(data, 9) = "SPROCESS=" Then
        'Request from client to start process
        Dim xProcess As String, Equal As Integer
        Equal = InStr(data, "=")
        xProcess = Right(data, Len(data) - Equal)
        Call StartProcess(xProcess, Server(Index))
    Exit Sub
    End If
    
     If Left(data, 9) = "TPROCESS=" Then
        'Request from client to terminate process
        Dim xProcessT As String, EqualSign As Integer
        EqualSign = InStr(data, "=")
        xProcessT = Right(data, Len(data) - EqualSign)
        Call TerminateRunningProcess(xProcessT, Server(Index))
    Exit Sub
    End If
    
    If Left(data, 3) = "DIR" Then '- Request for DIRECTORY listing
        Dim fsoObj As New FileSystemObject
        di = InStr(data, "DIR")
        Fname = Right(data, Len(data) - 4)

        'Make sure folder exists
        If fsoObj.FolderExists(Fname) = False Then
            Server(Index).SendData ("DIR->NOTFOUND")
            Set fsoObj = Nothing
            Exit Sub
        End If
        
        'Make sure the folder is listed in the shared folders
        If Not IsValidSharedFolder(CStr(Fname)) Then
            Server(Index).SendData ("NOT_SHARED")
            Exit Sub
        End If
        
        'Create data packet that represents the file(s) and folder(s)
        'inside the requested directory, and send it to the client
        buff = DirectoryToString(CStr(Fname))
        buff = "DIR->" & Fname & "|" & buff
        Server(Index).SendData (buff)
    
        sOutput "DIR '" & Fname & "' from IP '" & Server(Index).RemoteHostIP & "'"
    End If
    
    If data = "SHAREDFOLDERS" Then '- Request for shared folder list from client
        'Open the shared folders configuration file, and send the list
        'to the client
        i = FreeFile
        Open App.Path + "\SD.DLL" For Input As #i
            Do Until EOF(i):
                DoEvents
                Line Input #i, fldr
                buff = buff & fldr & "|"
            Loop
        Close #i
            Server(Index).SendData ("SF->" & buff)
    End If
End Sub


Private Sub ServerOutputHistory_MENU_Click()
DisplayLogFile "Output"
End Sub

Private Sub StartServer_MNU_Click()
If sEnabled = True Then MsgBox "Server already started.", 16, "Error": Exit Sub
Call EnableServer(True)
End Sub

Private Sub TerminateandExit_MNU_Click()
RetVal = MsgBox("Are you sure you wish to exit Andromeda RFS? Any active connections will be broken!", 36, "Really Exit?")

Select Case RetVal
    Case vbYes:
        End
End Select
End Sub

Private Sub TerminateServer_MNU_Click()
If sEnabled = False Then MsgBox "Server already stopped.", 16, "Error": Exit Sub
Call EnableServer(False)
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
    'Displays the number of active socket connections
    'to the server
    lblConnections.Caption = ttlLogins
End Sub

Private Sub TimerUptime_Timer()
'This timer increments the seconds, minutes and hours
'that the server has been running

firstcolon = InStr(txtElapsed.Caption, ":")
hourz = Left(txtElapsed.Caption, firstcolon - 1)
nextstring = Right(txtElapsed.Caption, Len(txtElapsed.Caption) - firstcolon)
colon = InStr(nextstring, ":")
minutez = Left(nextstring, colon - 1)
nextstring = Right(nextstring, Len(nextstring) - colon)
colon = InStr(nextstring, ":")
seconds = Right(nextstring, Len(nextstring) - colon)



If seconds = 59 Then
seconds = "00"
If minutez = 59 Then
minutez = "00"
hourz = hourz + 1
End If
minutez = minutez + 1
End If

seconds = seconds + 1
If Len(seconds) = 1 Then seconds = "0" & seconds
If Len(minutez) = 1 Then minutez = "0" & minutez
txtElapsed.Caption = hourz & ":" & minutez & ":" & seconds

End Sub

Private Sub tmrProcesses_Timer()
'Refresh running process list
KillApp "none", lstProcesses
End Sub


Private Sub TProcess_MENU_Click()
RetVal = MsgBox("Are you sure you wish to exit Andromeda RFS? Any active connections will be broken!", 36, "Really Exit?")

Select Case RetVal
    Case vbYes:
        End
End Select
End Sub

Private Sub TransferLog_MENU_Click()
Call DisplayLogFile("FileTransfer")
End Sub

Private Sub UDP_Close(Index As Integer)
frmMain.UDP(Index).Close
End Sub

Private Sub UDP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String
UDP(Index).GetData data, vbString, bytesTotal

 If Left(data, 3) = "GET" Then '- Request for file transfer
        
        ge = InStr(data, "GET")
        
        rest = Right(data, Len(data) - 4)
        
        Col = InStr(data, ":")
        
        mas = Right(data, Len(data) - Col)
        
        Col = InStr(mas, ":")
        
        data = Right(data, Len(data) - 4)
        
        f = Left(data, Col + 1)
        
        ip = Right(mas, Len(mas) - Col)
        sOutput "Request for '" & f & "' from IP '" & UDP(Index).RemoteHostIP & "'"
      
        SendFileToClient CStr(f), CStr(ip), UDP(Index)
    End If
    
End Sub

Private Sub UP_Close(Index As Integer)
UP(Index).Close
End Sub

Private Sub UP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
    upcount = UP.Count
    Load UP(upcount)
    UP(upcount).Accept requestID
End If

End Sub

Private Sub UP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String, FileSize As Long, Percent As Long

'On Error GoTo ErrorHandle

Call UP(Index).GetData(data, , bytesTotal)

    
    If Left(data, 5) = "FILE=" Then 'Received file upload confirmation from
                                    'client... separate data, and set variables
        
     '   temp$ = Right(Data, Len(Data) - 5)
     '   slash = FindReverse(temp$, "\")
     '   ParentFolder$ = Left(temp$, slash)
     '   'Debug.Print Data
     '   If Exists(ParentFolder$) = False Then
     '       MkDir (ParentFolder$)
     '   End If
        
        Dim folders2create As New Collection
       Dim objFso As New FileSystemObject
        data = Right(data, Len(data) - 5)
        
        colon = InStr(data, ":")
        
        nextstring = Right(data, Len(data) - colon)
        
        realcolon = InStr(nextstring, ":") + 2
        
        FileSize1 = Right(data, Len(data) - realcolon)
        
        fileName = Left(data, realcolon - 1)
        
        FileTransferAdd fileName, FileSize1, UP(Index).RemoteHostIP, "" 'Add item to list for file transfers
       
        pf = objFso.GetParentFolderName(fileName)
            Do While pf <> "": DoEvents
                If objFso.FolderExists(pf) = False Then
                    folders2create.Add pf
                End If
                pf = objFso.GetParentFolderName(pf)
            Loop
            
        'Create folders (if needed)
        On Error Resume Next
        For X = folders2create.Count To 1 Step -1
            MkDir folders2create.Item(X)
        Next X
        
        Set folders2create = Nothing
        'Delete the file
        If Exists(fileName) Then Kill fileName
        
        'Open the file so that packets received can be directly
        'written to the already open disk file
        fileNum = FreeFile()
        i = FreeFile
        Open fileName For Binary Access Write As #fileNum
        
        If FileSize1 = 0 Then
            'If the file size is 0 bytes, just close the file
            'and tell the client it's done receiving the file
            Close #fileNum
            Call frmMain.UP(Index).SendData("FILEDONE")
            sOutput "Received '" & fileName & "' (" & FileSize1 & " bytes) from IP '" & UP(Index).RemoteHostIP & "'"
            Exit Sub
        End If
            
        'Inform the client that it can start sending
        'data packets (the default is 2048 bytes)
        Call frmMain.UP(Index).SendData("BEGIN")
        Exit Sub
    End If
    
    'Inform the client that the packet was received sucessfully
    frmMain.UP(Index).SendData ("OK")
    
    'Write the incoming data directly to the disk file
    Put #fileNum, , data
    DoEvents
    
    'If the size of the disk file matches the size as told
    'by the client, we are done receiving this file, so
    'close it and inform the client that the file was
    'received successfully
     If LOF(fileNum) = FileSize1 Then
        Close #fileNum
        Debug.Print "Closed file#: " & fileNum
        Call frmMain.UP(Index).SendData("FILEDONE")
        sOutput "Received '" & fileName & "' (" & FileSize1 & " bytes) from IP '" & UP(Index).RemoteHostIP & "'"
        
        'If logging is enabled in options, write this transfer to the log
        If GetSetting("Andromeda", "Settings", "WriteTransferLog") = "1" Then
            WriteLog App.Path + "\FTransfer.txt", "Received '" & fileName & "' (" & FileSize1 & " bytes) from IP '" & UP(Index).RemoteHostIP & "' Time/Date=" & Format(Now, "HH:MM:SS AM/PM - MM/DD/YYYY")
        End If
        
        fileNum = 0 'Set fileNum back to zero
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandle:
    sOutput ("Error in UP(" & Index & "): " & Err.Description & " #: " & Err.Number)

End Sub

Private Sub UP_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Winsock Error in frmMain.UP(" & Index & ")" & vbCrLf & vbCrLf & Err.Description, 16, "Winsock TCP/IP Error"
End Sub

Private Sub VLoginHistory_MENU_Click()
DisplayLogFile "Login"
End Sub

Private Sub WebSite_MENU_Click()
Call ShellExecute(Me.hwnd, "open", "http://www.induhviduals.com/andromeda", 0, 0, vbNormalFocus)
End Sub



Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Index = 0 Then
        intMax2 = intMax2 + 1
        Load UDP(intMax2)
        Load Server(intMax2)
        If Server(intMax2).State <> sckClosed Then Server(intMax2).Close
        Server(intMax2).Accept (requestID)
    End If
End Sub




Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbCritical
End Sub


