VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSharedFolders 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ndromeda - Manage Shared Directories"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSharedFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add..."
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton btnUnshare 
      Caption         =   "&Unshare"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSharedFolders.frx":014A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstDirectories 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
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
      Width           =   7450
   End
   Begin VB.Label Label1 
      Caption         =   "Your Shared Directories:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmSharedFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()

frmAddSharedDirectory.Show vbModal, Me




End Sub

Private Sub btnSave_Click()
SaveSharedDirectories
Unload Me
End Sub

Private Sub btnUnshare_Click()
lstDirectories.ListItems.Remove (lstDirectories.SelectedItem.Index)
End Sub

Private Sub Form_Load()
LoadSharedDirectories

End Sub
