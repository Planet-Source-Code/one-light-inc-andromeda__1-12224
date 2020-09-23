VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda -  Manage User Accounts"
   ClientHeight    =   4320
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
   Icon            =   "frmManageUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstbuffer 
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   6495
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add..."
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnModify 
      Caption         =   "&Modify "
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "&Delete "
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Password"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Last Login"
         Object.Width           =   5363
      EndProperty
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
      Width           =   7800
   End
End
Attribute VB_Name = "frmManageUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
frmCreateUser.Show vbModal, Me
End Sub

Private Sub btnClose_Click()
Unload Me
End Sub

Private Sub btnModify_Click()
If ListView1.SelectedItem.Text = "" Then Exit Sub
Call ModifyUser(ListView1.SelectedItem.Text)
End Sub

Private Sub Command1_Click()
Text2.Text = Encrypt(Text1.Text)
End Sub

Private Sub Command2_Click()
Text1.Text = Decrypt(Text2.Text)
End Sub

Private Sub btnRemove_Click()
If ListView1.SelectedItem.Text = "" Then Exit Sub
arse = MsgBox("Are you sure you want to permanantly delete the User '" & ListView1.SelectedItem.Text & "' ?", 36, "Confirm User Deletion")

Select Case arse
    Case vbYes
        Kill App.Path + "\" + ListView1.SelectedItem.Text + ".alf"
        LoadExistingUserInformation
        MsgBox "User sucessfully deleted.", 64, "Information"
    Case vbNo
    
End Select

End Sub

Private Sub Form_Load()
LoadExistingUserInformation

End Sub
