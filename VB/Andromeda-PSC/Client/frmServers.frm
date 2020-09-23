VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Andromeda Servers"
   ClientHeight    =   4470
   ClientLeft      =   3765
   ClientTop       =   3015
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   6975
   Begin VB.Frame ServerFrame 
      Height          =   3965
      Left            =   4080
      TabIndex        =   3
      Top             =   -90
      Width           =   2895
      Begin VB.TextBox txtInitRemDir 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Initial Remote Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Login Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Server IP Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   480
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
            Picture         =   "frmServers.frx":0442
            Key             =   "Server"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Servers 
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6376
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton RemoveServer 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton NewServer 
      Caption         =   "&New"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Connect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame TopFrame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   25
      Left            =   -240
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Frame BottomFrame 
      Height          =   25
      Left            =   -480
      TabIndex        =   1
      Top             =   3840
      Width           =   7335
   End
End
Attribute VB_Name = "frmServers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldCaption As String
Dim Connecting As Boolean
Dim Closing As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


Sub LoadServers()

Dim fs As New FileSystemObject

Path = App.Path + "\Servers\"
MyDir = Dir(Path, vbDirectory)
    Do While MyDir <> "": DoEvents
        If MyDir <> "." And MyDir <> ".." Then
            ext = fs.GetExtensionName(MyDir)
            fn = fs.GetFileName(MyDir)
            If ext = "rfs" Then
                labl$ = ReadINI("", "ServerLabel", Path + CStr(MyDir))
                Call Servers.ListItems.Add(, , Left(fn, Len(fn) - 4), , "Server")
            End If
        End If
        MyDir = Dir
    Loop
End Sub



Private Sub btnExit_Click()
    Closing = True
    Unload Me
    frmMain.Toolbar1.Buttons(1).Value = tbrUnpressed
End Sub



Private Sub Connect_Click()
If Connect.Caption = "&Connect" Then
    Connect.Caption = "&Cancel"
    If Servers.SelectedItem Is Nothing Then Exit Sub
    
    If frmMain.Winsock.State <> sckClosed Then
        frmMain.Winsock.Close
    End If
    
    frmMain.Winsock.RemoteHost = txtServerIP
    frmMain.Winsock.RemotePort = 6969
    frmMain.Winsock.Connect
    
    Me.MousePointer = 13
    frmMain.MousePointer = 13
    
    Do While frmMain.Winsock.State <> sckConnected: DoEvents
        If frmMain.Winsock.State = sckError Then GoTo skip
    Loop
    
    Connecting = False
    
ElseIf Connect.Caption = "&Cancel" Then
skip:
    Connect.Caption = "&Connect"
    frmMain.Winsock.Close
    Me.MousePointer = 0
    frmMain.MousePointer = 0
    Connecting = False
End If
End Sub





Private Sub Form_Load()
    Call CenterFormMDI(frmMain, Me) ': Me.Top = Me.Top - frmMain.Toolbar1.Height
    Call LoadServers
    Call Servers_ItemClick(Servers.SelectedItem)
    Closing = False
    
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Closing = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Toolbar1.Buttons(1).Value = tbrUnpressed
    Closing = True
End Sub

Private Sub NewServer_Click()
    frmNewServerWizard.Show
End Sub

Private Sub RemoveServer_Click()
    If Servers.SelectedItem Is Nothing Then Exit Sub
    ret = MsgBox("Are you sure you wish to remove " + Servers.SelectedItem.Text + " from your server list?", 36, "Remove Server?")
    If ret = vbNo Then Exit Sub
    On Error Resume Next
    ind = Servers.SelectedItem.index
    Kill App.Path + "\Servers\" + Servers.SelectedItem.Text + ".rfs"
    Servers.ListItems.Remove (ind)
End Sub

Private Sub Servers_AfterLabelEdit(Cancel As Integer, NewString As String)
    If IsValidFileName(NewString) = False Then
        MsgBox "Invalid server name. Cannot contain these characters: \/:*?*" & Chr(34) & "<>|", vbCritical, "Invalid Name"
        Cancel = True
        Exit Sub
    End If
    
    'rename the configuration file
    Dim fso As New FileSystemObject
    fso.GetFile(App.Path + "\Servers\" + OldCaption + ".rfs").Name = NewString & ".rfs"

End Sub

Private Sub Servers_BeforeLabelEdit(Cancel As Integer)
    If Connecting = True Then Cancel = True: Exit Sub
    OldCaption = Servers.SelectedItem.Text
End Sub

Private Sub Servers_DblClick()
    If Servers.SelectedItem Is Nothing Then Exit Sub
    Call Connect_Click
    Connect.SetFocus
    Connecting = True
End Sub

Private Sub Servers_ItemClick(ByVal item As MSComctlLib.ListItem)
On Error Resume Next
    Path$ = App.Path + "\Servers\" + item.Text + ".rfs"
    
    txtServerLabel = ReadINI("", "ServerLabel", Path$)
    txtServerIP = ReadINI("", "ServerIP", Path$)
    txtLogin = ReadINI("", "LoginName", Path$)
    txtPassword = ReadINI("", "Password", Path$)
    txtInitRemDir = ReadINI("", "InitRemDir", Path$)
End Sub






Private Sub txtInitRemDir_Change()
    Dim InitRemDir As String
    InitRemDir = txtInitRemDir
    
    If InitRemDir <> "" Then
        If Right(InitRemDir, 1) <> "\" Then InitRemDir = InitRemDir + "\"
    End If
    
    Path$ = App.Path + "\Servers\" + Servers.SelectedItem.Text + ".rfs"
    Call WriteINI("", "InitRemDir", InitRemDir, Path$)
End Sub

Private Sub txtLogin_Change()
    Path$ = App.Path + "\Servers\" + Servers.SelectedItem.Text + ".rfs"
    Call WriteINI("", "LoginName", txtLogin, Path$)
End Sub

Private Sub txtPassword_Change()
    Path$ = App.Path + "\Servers\" + Servers.SelectedItem.Text + ".rfs"
    Call WriteINI("", "Password", txtPassword, Path$)
End Sub


Private Sub txtServerIP_Change()
    Path$ = App.Path + "\Servers\" + Servers.SelectedItem.Text + ".rfs"
    Call WriteINI("", "ServerIP", txtServerIP, Path$)
End Sub






