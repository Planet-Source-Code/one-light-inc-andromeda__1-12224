VERSION 5.00
Begin VB.Form frmModifyUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Modify User Settings"
   ClientHeight    =   1680
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
   Icon            =   "frmModifyUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame frameUser 
      Caption         =   "User settings for: "
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -20
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmModifyUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
hoto = MsgBox("Are you sure you wish to make changes to the settings for '" & txtUser.Text & "' ?", 36, "Confirm User Config Modification")

Select Case hoto
    Case vbYes
    Call WriteEncryptedINI("Andromeda", "PW", txtPassword.Text, App.Path + "\" + txtUser.Text + ".alf")
        Dim AccessRights As String
    If optNormalUser.Value = True Then
        AccessRights = "NORMAL_ACCESS"
    End If
    
    If optFullAccess.Value = True Then
        AccessRights = "FULL_ACCESS"
    End If
    
    If optSuperUser.Value = True Then
        AccessRights = "SUPER_USER"
    End If
    
    Call WriteEncryptedINI("Andromeda", "USER_RIGHTS", AccessRights, App.Path + "\" + txtUser.Text + ".alf")
    
    Select Case chkLogging.Value
        Case True
            Call WriteEncryptedINI("Andromeda", "Log", "True", App.Path + "\" + txtUser.Text + ".alf")

        Case False
            Call WriteEncryptedINI("Andromeda", "Log", "False", App.Path + "\" + txtUser.Text + ".alf")

    End Select
        
    MsgBox "Settings for '" & txtUser.Text & "' have been sucessfully modified.", 64, "Information": Unload Me
    Case vbNo
    
End Select
End Sub

