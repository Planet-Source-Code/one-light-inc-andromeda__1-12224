VERSION 5.00
Begin VB.Form frmProcesses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Start/Terminate Process"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcesses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   5535
   End
   Begin VB.CommandButton btnSpawn 
      Caption         =   "&Spawn"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtNewProcess 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "Type the path to the executable file here"
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   5535
   End
   Begin VB.CommandButton btnTerminate 
      Caption         =   "&Terminate"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox lstServersProcesses 
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Spawn this new process:"
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
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Processes currently running on the server:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnRefresh_Click()
    lstServersProcesses.Clear
    Call GetRunningProcesses
End Sub

Private Sub btnSpawn_Click()
    If txtNewProcess.Text = "" Then MsgBox "Nothing entered for new process to spawn.", 16, "Error": Exit Sub
    frmMain.Winsock.SendData ("SPROCESS=" & txtNewProcess.Text)

End Sub

Private Sub btnTerminate_Click()
    If lstServersProcesses.ListIndex = 0 Then MsgBox "No process selected.", 16, "Error": Exit Sub
    frmMain.Winsock.SendData ("TPROCESS=" & frmProcesses.lstServersProcesses.List(frmProcesses.lstServersProcesses.ListIndex))
End Sub

Private Sub Form_Load()
    Call GetRunningProcesses
    
End Sub

Private Sub txtNewProcess_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnSpawn_Click
    End If
End Sub


