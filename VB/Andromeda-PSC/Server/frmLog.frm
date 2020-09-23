VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Login File"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox WhichLog 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Erase"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtLogin 
      Height          =   4575
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -20
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label lblFileSize 
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Log File Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Me.Hide
End Sub

Private Sub btnDelete_Click()

RetVal = MsgBox("Are you sure you want to delete the Andromeda Log File?", 36, "Confirm Erase")

Select Case RetVal
    Case vbYes
        i = FreeFile
        Select Case WhichLog
            Case "Login":
                 Open App.Path + "\Log.txt" For Output As #i
                    Print ""
                 Close #i
            Case "FileTransfer":
                 Open App.Path + "\FTransfer.txt" For Output As #i
                    Print ""
                 Close #i
            Case "Output":
                Open App.Path + "\Output.txt" For Output As #i
                    Print ""
                Close #i
        End Select
        txtLogin.Text = ""
        lblFileSize = "0 bytes"
        MsgBox "The Log File was deleted.", 64, "Deleted Log File"
End Select

End Sub

Private Sub btnPrint_Click()
If txtLogin.Text = "" Then MsgBox "There is no data to print.", 16, "Error": Exit Sub
On Error GoTo err_
With CommonDialog1

    .ShowPrinter
    .CancelError = True
End With

    Printer.FontSize = 10
    Printer.Print "Access Log File for Andromeda RFS" & vbCrLf & "Created " & Format(Now, "mm/dd/yy - HH:MM:SS AM/PM") & vbCrLf & vbCrLf & txtLogin.Text
    Printer.EndDoc
Exit Sub
err_:
End Sub


