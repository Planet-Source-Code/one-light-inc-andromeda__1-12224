VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   " Andromeda RFS (Client)"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1320
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox StatusWindow 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1815
      ScaleWidth      =   10215
      TabIndex        =   2
      Top             =   5520
      Width           =   10215
      Begin RichTextLib.RichTextBox txtStatus 
         Height          =   1455
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":220E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.Toolbar CloseStatus 
         Height          =   330
         Left            =   9840
         TabIndex        =   4
         Top             =   25
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7335
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   450
            MinWidth        =   450
            Picture         =   "frmMain.frx":4F04
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5640
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Show Server List"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "File Transfer"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences..."
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewServers 
         Caption         =   "View Server List"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuNewServerWiz 
         Caption         =   "New Server &Wizard"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Visible         =   0   'False
      Begin VB.Menu Processes_MNU 
         Caption         =   "Spawn/Terminate Server Processes"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTransfer 
         Caption         =   "File Transfer"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuStatusWindow 
         Caption         =   "Status"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Windows"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuFileTransfer_Click()
    frmFileView.Show
    frmFileView.Caption = "File Transfer - " + CurrentServer.ServerLabel
End Sub

Private Sub mnuPreferences_Click()
    frmOptions.Show vbApplicationModal
End Sub

'Implements iSubClass
Private Sub mnuServer_Click()
    If CurrentServer.ServerIP <> "" Then
        mnuConnect.Caption = "&Connect to " + CurrentServer.ServerLabel
    Else
        mnuConnect.Caption = "&Connect"
    End If
End Sub

Private Sub CloseStatus_ButtonClick(ByVal Button As MSComctlLib.Button)
    StatusWindow.Visible = False
    mnuStatusWindow.Checked = False
End Sub

Private Sub MDIForm_Load()
Me.WindowState = vbMaximized
If ShowServerList = 1 Then
    frmServers.Show
    Toolbar1.Buttons(1).Value = tbrPressed
End If

'frmFileView.Show
End Sub


Private Sub MDIForm_Resize()
On Error Resume Next
txtStatus.Width = Me.Width - 100
CloseStatus.Left = Me.Width - CloseStatus.Width - 100
End Sub


Private Sub mnuArrange_Click()
Me.Arrange (vbArrangeIcons)
End Sub

Private Sub mnuCascade_Click()
Me.Arrange (vbCascade)
End Sub

Private Sub mnuConnect_Click()
    If Winsock.State <> sckClosed Then
        Winsock.Close
        Do While Winsock.State <> sckClosed: DoEvents
        Loop
    End If
    
    Winsock.RemoteHost = CurrentServer.ServerIP
    Winsock.RemotePort = 6969
    Winsock.Connect
    
    Do While Winsock.State <> 7: DoEvents
        If Winsock.State = sckError Then Exit Sub
    Loop
    
    Winsock.SendData ("BEGIN_LOGIN")
    
End Sub

Private Sub mnuDisconnect_Click()
    Call Winsock.Close
    While Winsock.State <> sckClosed: DoEvents: Wend
    Status.Panels(1).Picture = ImageList1.ListImages(2).Picture
    mnuDisconnect.Enabled = False
    mnuConnect.Enabled = True
    txtStatus.SelColor = vbRed
    txtStatus.SelBold = True
    txtStatus.SelText = "Disconnected from server at " + Format(Time, "h:mm AM/PM") + vbCrLf
    txtStatus.SelBold = False
    mnuActions.Visible = False
    Unload frmFileView
    Unload frmProcesses
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuNewServerWiz_Click()
    frmNewServerWizard.Show
End Sub

Private Sub mnuStatusWindow_Click()
    With mnuStatusWindow
        If .Checked = True Then
            StatusWindow.Visible = False
            .Checked = False
        Else
            StatusWindow.Visible = True
            .Checked = True
        End If
    End With
End Sub

Private Sub mnuTileHor_Click()
Me.Arrange (vbTileHorizontal)
End Sub

Private Sub mnuTileVert_Click()
Me.Arrange (vbTileVertical)
End Sub

Private Sub mnuViewServers_Click()
    Toolbar1.Buttons(1).Value = tbrPressed
    frmServers.Show vbApplicationModal
End Sub



Private Sub Processes_MNU_Click()
frmProcesses.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1
            'view server list
            Button.Value = tbrPressed
            frmServers.Show
            
        Case 3
            'view file transfer
            Button.Value = tbrPressed
            frmFileView.Show
    End Select
End Sub


Private Sub txtStatus_Change()
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

Private Sub Winsock_Close()
    Status.Panels(1).Picture = ImageList1.ListImages(2).Picture
    
    frmMain.Toolbar1.Buttons(3).Enabled = False
    
    mnuDisconnect.Enabled = False
    mnuConnect.Enabled = True
    
    txtStatus.SelColor = vbRed
    txtStatus.SelBold = True
    txtStatus.SelText = "Connection Lost! (" + Format(Time, "h:mm AM/PM") + ")" + vbCrLf
    txtStatus.SelBold = False
    
    Winsock.Close
    frmServers.MousePointer = 0
    frmServers.Connect.Caption = "&Connect"
    Me.MousePointer = 0
    mnuActions.Visible = False
    Unload frmFileView
    Unload frmProcesses
    
    If AutoReconnect = 1 Then
        Call mnuConnect_Click
    End If
End Sub

Private Sub Winsock_Connect()
    CurrentServer.ServerLabel = frmServers.Servers.SelectedItem.Text
    CurrentServer.ServerIP = Winsock.RemoteHostIP
    CurrentServer.Login = frmServers.txtLogin
    CurrentServer.Password = frmServers.txtPassword
    CurrentServer.InitDir = frmServers.txtInitRemDir
    
    txtStatus.SelColor = vbBlue
    txtStatus.SelBold = True
    txtStatus.SelText = "Connected to " & CurrentServer.ServerLabel & " (" & Winsock.RemoteHostIP & ") on port " & Winsock.RemotePort & vbCrLf
    txtStatus.SelBold = False
    txtStatus.SelText = "Waiting For Login Request..." & vbCrLf
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    'This is where all the commands are handled
    
    Dim Data As String
    Call Winsock.GetData(Data, , bytesTotal)
    
    If Left(Data, 6) = "ERROR:" Then
        'Server Error
        Data = Right(Data, Len(Data) - 6)
        Call ServerError(Data)
        Exit Sub
    End If
    
    If Data = "LOGIN" Then
        txtStatus.SelColor = vbBlue
        txtStatus.SelText = "Received Login Request From Server" + vbCrLf
        Call Winsock.SendData("LOGIN=" + CurrentServer.Login + ":" + CurrentServer.Password)
        txtStatus.SelText = "Sent Login to Server...Waiting for validation" + vbCrLf
        Exit Sub
    End If
    
    If Data = "INVALID_LOGIN" Then
        txtStatus.SelColor = vbRed
        txtStatus.SelText = "Invalid Login Name!" + vbCrLf
        Winsock.Close
        frmServers.MousePointer = 0
        frmServers.Connect.Caption = "&Connect"
        Me.MousePointer = 0
        Exit Sub
    End If
    
    If Data = "INVALID_PASSWORD" Then
        txtStatus.SelColor = vbRed
        txtStatus.SelText = "Invalid Password for " + CurrentServer.Login + "!" + vbCrLf
        Winsock.Close
        frmServers.MousePointer = 0
        frmServers.Connect.Caption = "&Connect"
        Me.MousePointer = 0
        Exit Sub
    End If
    
    If Data = "WELCOME" Then
        txtStatus.SelColor = vbBlue
        txtStatus.SelBold = True
        txtStatus.SelText = "Logged in as " + CurrentServer.Login + " at " + Format(Time, "h:mm AM/PM") + vbCrLf
        txtStatus.SelBold = False
        Status.Panels(1).Picture = ImageList1.ListImages(1).Picture
        mnuDisconnect.Enabled = True
        mnuConnect.Enabled = False
        Unload frmServers
        frmMain.Toolbar1.Buttons(1).Value = tbrUnpressed
        frmMain.Toolbar1.Buttons(3).Enabled = True
        Me.MousePointer = 0
        mnuActions.Visible = True
        
        If AutoshowFileTransfer = 1 Then
            frmFileView.Show
        End If
        Exit Sub
    End If
    
    If Left(Data, 4) = "SF->" Then
        Call ShowSharedFolders(Right(Data, Len(Data) - 4))
        Exit Sub
    End If
    
    If Left(Data, 14) = "DIR_CONTENTS->" Then
        MsgBox Data
        'used to get contents of a folder (for downloading)
        Call frmFileView.StringToQue(Mid(Data, 15, Len(Data) - 15))
        WaitingForContents = False
        Exit Sub
    End If
        
    If Left(Data, 5) = "DIR->" Then
        If Data = "DIR->NOTFOUND" Then
            MsgBox "That directory does not exist!", vbCritical, "Server Path Not Found"
            For x = 1 To frmFileView.ServerDrives.ComboItems.Count
                If LCase(frmFileView.ServerDrives.ComboItems(x)) = LCase(ServerPath) Then
                    frmFileView.ServerDrives.ComboItems(x).Selected = True
                    GoTo skipit
                End If
            Next x
            GoTo skipit
        End If
        
        Dim NewDir As String, d As String
        pipe = InStr(Data, "|")
        NewDir = Mid(Data, 6, pipe - 6)
        d = Right(Data, Len(Data) - pipe)
        
        Call StringToDir(frmFileView.ServerFileList, d)
        ServerPath = NewDir
        
        For x = 1 To frmFileView.ServerDrives.ComboItems.Count
            If LCase(frmFileView.ServerDrives.ComboItems(x).Text) = LCase(ServerPath) Then
                frmFileView.ServerDrives.ComboItems(x).Selected = True
                GoTo skipit
            End If
        Next x
        
        Call frmFileView.ServerDrives.ComboItems.Add(1, , ServerPath, "folder")
        frmFileView.ServerDrives.ComboItems(1).Selected = True
        
skipit:
        frmFileView.ServerFileList.MousePointer = 0
        frmFileView.Status.Panels(3).Text = UCase(ServerPath)
        Exit Sub
    End If
    
    
    If Left(Data, 10) = "NOT_SHARED" Then
        For x = 1 To frmFileView.ServerDrives.ComboItems.Count
            If UCase(frmFileView.ServerDrives.ComboItems(x)) = UCase(ServerPath) Then
                frmFileView.ServerDrives.ComboItems(x).Selected = True
                Exit For
            End If
        Next x
        MsgBox "That folder is not shared!", vbCritical, "Not Shared"
        Exit Sub
    End If
    
    
    If Data = "RENAMED" Or _
       Data = "DELETED" Then
        Call Winsock.SendData("DIR " + ServerPath)
        frmFileView.ServerFileList.MousePointer = 13
        Exit Sub
    End If
    
    
    If Data = "MOVED" Then
        WaitForMove = False
        Exit Sub
    End If
    
    If Data = "NOTMOVED" Then
        MsgBox "An error occured while moving file(s).", vbCritical, "Error"
        WaitForMove = False
        Exit Sub
    End If
    
    
    If Data = "CREATED" Then
        WaitForFolder = False
        Exit Sub
    End If
    
    If Data = "NOTCREATED" Then
        WaitForFolder = False
        MsgBox "An error occurred while creating a new folder!", vbCritical, "Error"
        Exit Sub
    End If
    
    
    If Left(Data, 11) = "TERMINATED=" Then
        strProcess = Right(Data, Len(Data) - 11)
        MsgBox "Process '" & strProcess & "' was terminated successfully.", 64, "Process Terminated"
        Exit Sub
    End If
    
    If Left(Data, 8) = "STARTED=" Then
        strProcess = Right(Data, Len(Data) - 8)
        MsgBox "Process '" & strProcess & "' was started successfully.", 64, "Process started"
        Exit Sub
    End If
    
    
    If Left(Data, 11) = "PROCESSES->" Then
        'We got the running processes back from the server
        
        Dim strData As String
        strData = Right(Data, Len(Data) - InStr(Data, "PROCESSES->") - 10)
        Call InitializeProcessList(strData)
    End If
    
End Sub


Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    frmServers.MousePointer = 0
    Me.MousePointer = 0
    frmServers.Connect.Caption = "&Connect"
    
    Select Case Number
        Case 10061 'could not find server
            MsgBox "Server is not online!", vbCritical, "Error!"
            Exit Sub
    End Select
    MsgBox Description, vbCritical, "Error!"
End Sub


