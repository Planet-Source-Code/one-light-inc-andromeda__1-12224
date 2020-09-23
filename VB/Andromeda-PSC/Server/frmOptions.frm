VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Options"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Height          =   30
      Left            =   120
      TabIndex        =   17
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Security"
      Height          =   3255
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkProcessToggle 
         Caption         =   "Allow Process Start/Stop"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtMessage 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox chkRename 
         Caption         =   "Allow users to rename files"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkMove 
         Caption         =   "Allow users to move files"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Allow users to delete files"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Message for invalid users:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Logging"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   3495
      Begin VB.CheckBox chkRecordFileTransfers 
         Caption         =   "Record file transfers (\FTransfer.txt)"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkLogLogins 
         Caption         =   "Save Logins to Log (\Log.txt)"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "General"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
      Begin VB.CheckBox chkStartWithWindows 
         Caption         =   "Start with Windows"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox chkMinimizeToTray 
         Caption         =   "&Minimize to tray"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkSplash 
         Caption         =   "&Show splash screen"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   640
         Width           =   1935
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -20
      TabIndex        =   0
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
'Save options to registry using SaveSetting
If chkLogLogins.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "WriteLog", "1"
Else
    SaveSetting "Andromeda", "Settings", "WriteLog", "0"
End If

If chkMinimizeToTray.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "MinimizeToTray", "1"
Else
    SaveSetting "Andromeda", "Settings", "MinimizeToTray", "0"
End If

If chkProcessToggle.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "AllowProcessToggle", "1"
Else
    SaveSetting "Andromeda", "Settings", "AllowProcessToggle", "0"
End If

If chkSplash.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "SplashScreen", "1"
Else
    SaveSetting "Andromeda", "Settings", "SplashScreen", "0"
End If

If chkRecordFileTransfers.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "WriteTransferLog", "1"
Else
    SaveSetting "Andromeda", "Settings", "WriteTransferLog", "0"
End If

If chkDelete.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "AllowDelete", "1"
Else
    SaveSetting "Andromeda", "Settings", "AllowDelete", "0"
End If

If chkRename.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "AllowRename", "1"
Else
    SaveSetting "Andromeda", "Settings", "AllowRename", "0"
End If

If chkMove.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "AllowMove", "1"
Else
    SaveSetting "Andromeda", "Settings", "AllowMove", "0"
End If

If chkStartWithWindows.Value = 1 Then
    SaveSetting "Andromeda", "Settings", "StartWithWindows", "1"
    WriteRegistry HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "AndromedaRFS", App.Path + "\" + App.EXEName + ".exe"
Else
    SaveSetting "Andromeda", "Settings", "StartWithWindows", "0"
    RemoveFromRegistry
End If

'Save invalid login message to file
i = FreeFile
Open App.Path + "\imessage.txt" For Output As #i
    Print #i, txtMessage.Text
Close #i
Me.Hide
End Sub

Private Sub Form_Load()
'Load options from registry
If GetSetting("Andromeda", "Settings", "WriteLog") = "1" Then
    chkLogLogins.Value = 1
Else
    chkLogLogins.Value = 0
End If

If GetSetting("Andromeda", "Settings", "MinimizeToTray") = "1" Then
    chkMinimizeToTray.Value = 1
Else
    chkMinimizeToTray.Value = 0
End If

If GetSetting("Andromeda", "Settings", "SplashScreen") = "1" Then
    chkSplash.Value = 1
Else
    chkSplash.Value = 0
End If

If GetSetting("Andromeda", "Settings", "WriteTransferLog") = "1" Then
    chkRecordFileTransfers.Value = 1
Else
    chkRecordFileTransfers.Value = 0
End If

If GetSetting("Andromeda", "Settings", "AllowDelete") = "1" Then
    chkDelete.Value = 1
Else
    chkDelete.Value = 0
End If

If GetSetting("Andromeda", "Settings", "AllowRename") = "1" Then
    chkRename.Value = 1
Else
    chkRename.Value = 0
End If

If GetSetting("Andromeda", "Settings", "AllowMove") = "1" Then
    chkMove.Value = 1
Else
    chkMove.Value = 0
End If

If GetSetting("Andromeda", "Settings", "StartWithWindows") = "1" Then
    chkStartWithWindows.Value = 1
Else
    chkStartWithWindows.Value = 0
End If

If GetSetting("Andromeda", "Settings", "AllowProcessToggle") = "1" Then
    chkProcessToggle.Value = 1
Else
    chkProcessToggle.Value = 0
End If
txtMessage.Text = InvalidMessage

End Sub


