VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmUpload 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connecting..."
   ClientHeight    =   3855
   ClientLeft      =   1410
   ClientTop       =   1860
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CloseWhenDone 
      Caption         =   "Close this window when upload completes"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   3840
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
      Begin VB.Timer Counter 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4440
         Top             =   1200
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Kps 
         Caption         =   "Transfer Rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label NumFiles 
         Caption         =   "Connecting to server..."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label BytesSent 
         Caption         =   "Bytes sent"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label TimeElapsed 
         Caption         =   "Time Elapsed:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   4935
      End
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1720
      _Version        =   393216
      FullWidth       =   345
      FullHeight      =   65
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Seconds As Integer
Dim bps As Long

Private Sub btnCancel_Click()
    Winsock.Close
    CancelUpload = True
    Unload Me
End Sub


Private Sub Counter_Timer()
    Seconds = Seconds + 1
    TimeElapsed.Caption = "Time Elapsed: " & FormatTime(CStr(Seconds))
    
    Kps.Caption = "Transfer Rate: " & FormatNumber((bps / 1000), 1) & " Kb/sec"
    bps = 0
End Sub

Private Sub Form_Load()
    Animation1.Open (App.Path + "\images\filemove.avi")
    Animation1.Play
End Sub




Private Sub Form_Unload(Cancel As Integer)
    CancelUpload = True
End Sub


Private Sub Winsock_Close()
    Winsock.Close
    CancelUpload = True
End Sub

Private Sub Winsock_Connect()
    Seconds = 0
    Me.Caption = "Uploading"
End Sub


Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    Call Winsock.GetData(Data, , bytesTotal)
    
    If Data = "BEGIN" Then
        StartSending = True
        Exit Sub
    End If
    
    If Data = "OK" Then
        WaitForServerRecieve = False
    End If
    
    If Data = "FILEDONE" Then
        FileDone = True
        Exit Sub
    End If
End Sub







Private Sub Winsock_SendProgress(ByVal BytesSent As Long, ByVal bytesRemaining As Long)
    bps = bps + BytesSent
End Sub


