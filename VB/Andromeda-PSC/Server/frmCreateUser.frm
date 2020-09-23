VERSION 5.00
Begin VB.Form frmCreateUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Create New User"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreateUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton btnCreateUser 
      Caption         =   "Create &User"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtPassword2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtPassword1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtLoginName 
      Height          =   285
      Left            =   1920
      MaxLength       =   16
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   -20
      TabIndex        =   0
      Top             =   0
      Width           =   5200
   End
   Begin VB.Frame Frame3 
      Caption         =   "User Name and Password"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.Label Label6 
         Caption         =   "User Name:"
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
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Desired Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Password Again:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCreateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnCreateUser_Click()
If Len(txtLoginName.Text) < 3 Then MsgBox "The new User Name must be atleast 3 characters long.", 16, "Error: User Name Invalid": Exit Sub
If UCase(txtPassword1.Text) <> UCase(txtPassword2.Text) Then MsgBox "The passwords entered in both provided fields must match.", 16, "Error: Passwords do not match": Exit Sub
'if we got this far, the username and passwords are acceptable
Dim UserRights As String, Logging As String

'Write user file
i = FreeFile
fileName$ = App.Path + "\" & txtLoginName.Text & ".alf"

Open fileName$ For Output As #i
    Print #i, "[Andromeda]"
    Print #i, "UserName=" & Encrypt(txtLoginName.Text & vbCrLf)
    Print #i, "PW=" & Encrypt(txtPassword1.Text & vbCrLf)
    Print #i, "LastLogin="
Close #i
LoadExistingUserInformation
MsgBox "New Login Account: " & txtLoginName.Text & " created!", 64, "Information"
Unload Me
End Sub

Private Sub Form_Load()
'Me.Caption = "Create New Login Account"
End Sub
