VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Preferences"
   ClientHeight    =   2385
   ClientLeft      =   4950
   ClientTop       =   4365
   ClientWidth     =   4215
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4215
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame frameServers 
      Caption         =   "Servers:"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3975
      Begin VB.CheckBox chkAutoshowFileTransfer 
         Caption         =   "Show file transfer window after connecting"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CheckBox chkAutoReconnect 
         Caption         =   "Automatically reconnect if disconnected"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   3255
      End
   End
   Begin VB.Frame frameGeneral 
      Caption         =   "General:"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkShowServerList 
         Caption         =   "Show server list at startup"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkShowSplash 
         Caption         =   "Show splash screen at startup"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOk_Click()

    WriteINI "Andromeda", "AutoReconnect", chkAutoReconnect, DLL()
    WriteINI "Andromeda", "AutoshowFileTransfer", chkAutoshowFileTransfer, DLL()
    WriteINI "Andromeda", "ShowSplash", chkShowSplash, DLL()
    WriteINI "Andromeda", "ShowServerList", chkShowServerList, DLL()
    
    Unload Me
End Sub

Private Sub chkAutoReconnect_Click()
    AutoReconnect = chkAutoReconnect
End Sub

Private Sub chkAutoshowFileTransfer_Click()
    AutoshowFileTransfer = chkAutoshowFileTransfer
End Sub



Private Sub chkShowSplash_Click()
    ShowSplash = chkShowSplash
End Sub

Private Sub Form_Load()
    Call CenterFormMDI(frmMain, Me)
    chkAutoReconnect = AutoReconnect
    chkAutoshowFileTransfer = AutoshowFileTransfer
    chkShowSplash = ShowSplash
    chkShowServerList = ShowServerList
    
End Sub
