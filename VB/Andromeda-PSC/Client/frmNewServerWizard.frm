VERSION 5.00
Begin VB.Form frmNewServerWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Server Wizard"
   ClientHeight    =   2835
   ClientLeft      =   9735
   ClientTop       =   7470
   ClientWidth     =   4680
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
   ScaleHeight     =   2835
   ScaleWidth      =   4680
   Begin VB.Frame Step 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2055
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtServerLabel 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "To add a new server to your list, please fill out these fields and click next."
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   0
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmNewServerWizard.frx":0000
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Enter the IP address for the server:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Type a label for the new server below:"
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.Frame Step 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2055
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtInitRemDir 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "Enter Initial Remote Directory:"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.Frame Step 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2175
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtLoginName 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Your Password:"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Your Login Name:"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   5535
      End
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "<  &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next  >"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   -120
      TabIndex        =   2
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmNewServerWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentStep As Integer

Private Sub btnBack_Click()
    Select Case CurrentStep
        Case 2 'login / password
            btnBack.Enabled = False
            Step(1).ZOrder
            txtServerLabel.SetFocus
            CurrentStep = 1
        Case 3 'init remote dir
            btnNext.Caption = "&Next  >"
            Step(2).ZOrder
            txtLoginName.SetFocus
            CurrentStep = 2
    End Select
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnNext_Click()
    Select Case CurrentStep
        Case 1 'server label / server ip
            If txtServerLabel = "" Or txtServerIP = "" Then
                MsgBox "You must fill out all fields before proceeding.", vbCritical, "Empty Field(s)"
                Exit Sub
            End If
            btnBack.Enabled = True
            Step(2).ZOrder
            txtLoginName.SetFocus
            CurrentStep = 2
            
        Case 2 'login / password
            If txtLoginName = "" Or txtPassword = "" Then
                MsgBox "You must fill out all fields before proceeding.", vbCritical, "Empty Field(s)"
                Exit Sub
            End If
            btnNext.Caption = "&Finish"
            Step(3).ZOrder
            txtInitRemDir.SetFocus
            CurrentStep = 3
            
        Case 3 'init remote dir
            If txtInitRemDir <> "" Then
                If Right(txtInitRemDir, 1) <> "\" Then txtInitRemDir = txtInitRemDir + "\"
            End If
            
            On Error Resume Next
            If objFso.FolderExists(App.Path + "\Servers") = False Then
                MkDir App.Path + "\Servers"
            End If
            
            Path$ = App.Path + "\Servers\" + txtServerLabel + ".rfs"
            Call WriteINI("", "ServerIP", txtServerIP, Path$)
            Call WriteINI("", "LoginName", txtLoginName, Path$)
            Call WriteINI("", "Password", txtPassword, Path$)
            Call WriteINI("", "InitRemDir", txtInitRemDir, Path$)
            
            
            Call frmServers.Servers.ListItems.Add(, , txtServerLabel, , "Server")
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Step(1).ZOrder
    CurrentStep = 1
End Sub


