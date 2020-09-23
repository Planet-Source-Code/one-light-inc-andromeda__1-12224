VERSION 5.00
Begin VB.Form frmAddSharedDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Add Shared Directory"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddSharedDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -20
      TabIndex        =   0
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmAddSharedDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOk_Click()
Dim j As ListItem
Dim fileName As String
fileName = Dir1.Path
If Right(fileName, 1) <> "\" Then fileName = fileName + "\"
Set j = frmSharedFolders.lstDirectories.ListItems.Add(, , fileName, , 1)
Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo err_
Dir1.Path = Drive1.Drive
Exit Sub
err_:
MsgBox "Device not ready!", 16, "Error"
End Sub


