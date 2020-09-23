VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ndromeda - Error"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H8000000F&
      Height          =   975
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "An error has occurred. Below is the server's response:"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmError.frx":014A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    Beep
End Sub


