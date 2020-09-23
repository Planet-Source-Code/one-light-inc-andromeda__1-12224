VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connecting..."
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5415
   Begin VB.CheckBox CloseWhenDone 
      Caption         =   "Close this window when download completes"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   2280
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
      TabIndex        =   1
      Top             =   3360
      Width           =   975
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
   Begin VB.Frame SingleFile 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5175
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
      Begin VB.Label NumFiles 
         Caption         =   "Connecting to server..."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label TimeRemaining 
         Caption         =   "Time Remaining:                      "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label TimeElapsed 
         Caption         =   "Time Elapsed:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label bytes 
         Caption         =   "Bytes copied"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label Kps 
         Caption         =   "Transfer Rate:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bps As Long



Private Sub btnCancel_Click()
    Winsock.Close
    CancelDownload = True
    On Error Resume Next
    Close FileNum
    Unload Me
End Sub

Private Sub Form_Load()
    Animation1.Open (App.Path + "\images\filemove.avi")
    Animation1.Play
End Sub







Private Sub Form_Unload(Cancel As Integer)
    CancelDownload = True
End Sub


Private Sub Timer1_Timer()
    Kps.Caption = "Transfer Rate: " & FormatNumber((bps / 1000), 1) & " Kb/sec"
    If bps <> 0 Then 'if bps is 0 the next line will cause error 'division by 0'
        TimeRemaining.Caption = "Time Remaining: " & FormatTime(FormatNumber(((FileSize - LOF(FileNum)) / bps), 0))
    End If
    bps = 0
End Sub



Private Sub Winsock_Close()
    Winsock.Close
    CancelDownload = True
    Unload Me
    
    On Error Resume Next
    Close FileNum
End Sub



Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    Winsock.Close
    Winsock.Accept (requestID)
End Sub


Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    SingleFile.Caption = "Copying " + CurrentFile + " from " + CurrentServer.ServerLabel
    
    Dim Data As String
    Dim Percent As Integer
    
    Call Winsock.GetData(Data, , bytesTotal)
    
    If Left(Data, 9) = "FILESIZE=" Then
        FileSize = Right(Data, Len(Data) - 9)
        If FileSize = 0 Then
            Close #FileNum
            fComplete = True
            Exit Sub
        End If
        Call frmMain.Winsock.SendData("BEGIN")
        Exit Sub
    End If
    
        
    bps = bps + bytesTotal
  
    Put #FileNum, , Data
    
    If LOF(FileNum) = FileSize Then
        Close #FileNum
        fComplete = True
        Exit Sub
    End If
    
    On Error Resume Next
    Percent = LOF(FileNum) / FileSize * 100
    Progress.Value = Percent
    Me.Caption = Percent & "% of " & CurrentFile & " downloaded"
    bytes.Caption = FormatFileSize(LOF(FileNum)) & " of " & FormatFileSize(FileSize) & " copied"
    
    
End Sub

