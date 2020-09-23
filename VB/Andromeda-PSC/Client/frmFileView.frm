VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileView 
   Caption         =   "File Transfer"
   ClientHeight    =   6315
   ClientLeft      =   3945
   ClientTop       =   3150
   ClientWidth     =   8670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileView.frx":0000
   LinkTopic       =   "Form2"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8670
   Begin MSComctlLib.ImageList SmallIcons 
      Left            =   5880
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":27A2
            Key             =   "picture"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":35F6
            Key             =   "cd-rom"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":5DAA
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":855E
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":AD12
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":D4C6
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":E74A
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":F9CE
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":12182
            Key             =   "font"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":125D6
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":14D8A
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1600E
            Key             =   "media"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":17292
            Key             =   "misc"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":17D5E
            Key             =   "rtf"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":181B2
            Key             =   "sound"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":19436
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1A6BA
            Key             =   "doc"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1B9C6
            Key             =   "zip"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList LargeIcons 
      Left            =   6480
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":1DB02
            Key             =   "cd-rom"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":202B6
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":22A6A
            Key             =   "desktop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2521E
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":279D2
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":28C56
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":29EDA
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2C68E
            Key             =   "font"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2CAE2
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":2F296
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3051A
            Key             =   "media"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3179E
            Key             =   "misc"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3226A
            Key             =   "rtf"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":326BE
            Key             =   "sound"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":33942
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":34BC6
            Key             =   "doc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":35ED2
            Key             =   "zip"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3800E
            Key             =   "picture"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":38E62
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":38FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3935A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":397AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":3990A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":39A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":39BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":39D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileView.frx":39E7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ServerFiles 
      Align           =   4  'Align Right
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   4335
      ScaleHeight     =   5940
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin MSComctlLib.ImageCombo ServerDrives 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "SmallIcons"
      End
      Begin MSComctlLib.ListView ServerFileList 
         Height          =   4815
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   8493
         SortKey         =   2
         View            =   2
         Arrange         =   2
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         _Version        =   393217
         Icons           =   "LargeIcons"
         SmallIcons      =   "SmallIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Go Up"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Download"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Large Icons"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Small Icons"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "List"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Details"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New Folder"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblServerFiles 
         Alignment       =   2  'Center
         BackColor       =   &H80000003&
         Caption         =   "Server's Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000013&
         Height          =   300
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4335
      End
   End
   Begin VB.PictureBox ClientFiles 
      Align           =   3  'Align Left
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6000
      Left            =   0
      ScaleHeight     =   5940
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   0
      Width           =   4215
      Begin MSComctlLib.ImageCombo ClientDrives 
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "SmallIcons"
      End
      Begin MSComctlLib.ListView ClientFileList 
         Height          =   4815
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8493
         SortKey         =   2
         View            =   2
         Arrange         =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         _Version        =   393217
         Icons           =   "LargeIcons"
         SmallIcons      =   "SmallIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Go Up"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Upload"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Large Icons"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Small Icons"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "List"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Details"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New Folder"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblClientFiles 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "My Files"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4215
      End
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   115
            MinWidth        =   115
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuLargeIcons 
         Caption         =   "Lar&ge Icons"
      End
      Begin VB.Menu mnuSmallIcons 
         Caption         =   "S&mall Icons"
      End
      Begin VB.Menu mnuList 
         Caption         =   "&List"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh Views"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuClient 
      Caption         =   "&Client"
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload Selected Item(s)"
         Shortcut        =   ^U
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClientMoveTo 
         Caption         =   "&Move to..."
      End
      Begin VB.Menu mnuRenameClientFile 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDeleteClientFile 
         Caption         =   "&Delete"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewClient 
         Caption         =   "&New"
         Begin VB.Menu mnuNewClientFolder 
            Caption         =   "&Folder"
         End
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download Selected Item(s)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServerMoveTo 
         Caption         =   "&Move to..."
      End
      Begin VB.Menu mnuRenameServerFile 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuDeleteServerFile 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecuteServerProgram 
         Caption         =   "E&xecute (on server)"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuNewFolder 
            Caption         =   "&Folder"
         End
      End
   End
End
Attribute VB_Name = "frmFileView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstRun As Boolean

Dim WithEvents Que As FileQue
Attribute Que.VB_VarHelpID = -1
Dim WithEvents UpQue As FileQue
Attribute UpQue.VB_VarHelpID = -1

Dim Focus As String

Private Source As String



Sub AddFileToQue(Path As String, DestPath As String)
    On Error Resume Next
    Que.AddFile (Path & ">" & DestPath)
End Sub

Sub AddFolderToQue(Path As String, DestPath As String, WhichQue As FileQue)
    On Error Resume Next
    
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    
    'This adds the files inside 'Path' (if any)
    mypath = Path
    myname = Dir(mypath)
    Do While myname <> "": DoEvents
        If myname <> "." And myname <> ".." Then
            If (GetAttr(mypath & myname) And vbDirectory) = vbDirectory Then GoTo next1
            WhichQue.AddFile (mypath & myname & ">" & DestPath)
        End If
next1:
        myname = Dir
    Loop
        
    Dim objDir1 As Folder
    Dim objDir2 As Folder
    Set objDir1 = objFso.GetFolder(Path)
    
    'This part adds all the files inside subfolders
    For Each objDir2 In objDir1.SubFolders
        parentfolders = Mid(Path, Len(ClientPath), Len(Path) - Len(ClientPath))
        Call AddFolderToQue(Path & objDir2.Name, ServerPath & parentfolders & "\" & objDir2.Name & "\", WhichQue) 'recursive call to this function
    Next objDir2
        
    Set objDir1 = Nothing
    Set objDir2 = Nothing
    
    
End Sub


Sub GetServerFilesFromFolder(Path As String)
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    frmMain.Winsock.SendData ("GET_DIR_CONTENTS=" & Path)
    
    WaitingForContents = True
    Do While WaitingForContents = True: DoEvents: Loop
End Sub

Sub RefreshClientView()
    Call ChowFromFolder(ClientFileList, ClientPath, "*.*")
End Sub

Private Sub RefreshServerView()
    Call frmMain.Winsock.SendData("DIR " + ServerPath)
    ServerFileList.MousePointer = 13
End Sub


Private Sub ShowFocus(str As String)
    Select Case str
        Case "Client"
            lblClientFiles.ForeColor = vbActiveTitleBarText
            lblClientFiles.BackColor = vbActiveTitleBar
            lblServerFiles.ForeColor = vbInactiveTitleBarText
            lblServerFiles.BackColor = vbInactiveTitleBar
            mnuServer.Enabled = False
            mnuClient.Enabled = True
            Focus = "Client"
        Case "Server"
            lblClientFiles.ForeColor = vbInactiveTitleBarText
            lblClientFiles.BackColor = vbInactiveTitleBar
            lblServerFiles.ForeColor = vbActiveTitleBarText
            lblServerFiles.BackColor = vbActiveTitleBar
            mnuClient.Enabled = False
            mnuServer.Enabled = True
            Focus = "Server"
    End Select
End Sub











Private Sub ClientDrives_Change()
    Call WriteINI("Andromeda", "LastClientFolder", ClientDrives.Text, DLL())
End Sub

Private Sub ClientDrives_Click()
    Call ChowFromFolder(ClientFileList, ClientDrives.Text, "*.*")
End Sub


Private Sub ClientDrives_GotFocus()
    ShowFocus "Client"
End Sub


Private Sub ClientDrives_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim fs As New FileSystemObject
        
        newpath = ClientDrives.Text
        
        If Right(newpath, 1) <> "\" Then newpath = newpath + "\"
        
        If fs.FolderExists(newpath) = True Then
            Call ChowFromFolder(ClientFileList, newpath, "*.*")
            For x = 1 To ClientDrives.ComboItems.Count
                If LCase(ClientDrives.ComboItems(x).Text) = LCase(newpath) Then
                    ClientDrives.ComboItems(x).Selected = True
                    Exit Sub
                End If
            Next x
            
            Dim itm As ComboItem
            Set itm = ClientDrives.ComboItems.Add(, , ClientDrives.Text, "folder")
            itm.Selected = True
            
        Else
            'folder doesnt exist
            Call RefreshClientView
        End If
    End If
End Sub

Private Sub ClientFileList_AfterLabelEdit(Cancel As Integer, NewString As String)
    If IsValidFileName(NewString) = False Then
        MsgBox "New file name contains illegal characters!", vbCritical, "Error"
        Cancel = 1
        Exit Sub
    End If
    
    If ClientFileList.SelectedItem.Icon = "folder" Then
        Call RenameFolder(ClientPath + ClientFileList.SelectedItem.Text, NewString)
    Else
        Call RenameFile(ClientPath + ClientFileList.SelectedItem.Text, NewString)
    End If
End Sub

Private Sub ClientFileList_DblClick()
    If ClientFileList.SelectedItem Is Nothing Then Exit Sub
    If ClientFileList.SelectedItem.Icon = "folder" Then
    'browse into this directory
        Call ChowFromFolder(ClientFileList, ClientPath + ClientFileList.SelectedItem.Text + "\", "*.*")
    Else
    'upload this file
        Call mnuUpload_Click
    End If
End Sub

Private Sub ClientFileList_GotFocus()
ShowFocus "Client"
End Sub






Private Sub ClientFileList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set ClientFileList.SelectedItem = Nothing
End Sub

Private Sub ClientFileList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        For x = 1 To ClientFileList.ListItems.Count
            If ClientFileList.ListItems(x).Selected = True Then n = n + 1
        Next x
        
        If n = 0 Then Exit Sub
        
        If n > 1 Then
            mnuRenameClientFile.Enabled = False
        Else
            mnuRenameClientFile.Enabled = True
        End If
        
        Call Me.PopupMenu(mnuClient)
    End If
End Sub

Private Sub ClientFileList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itm As ListItem
    Set itm = ClientFileList.HitTest(x, y)
    If itm Is Nothing Then Exit Sub
    
    If itm.Icon = "folder" Then
        'item(s) dropped in folder
        For x = 1 To Data.Files.Count
            Select Case Source
                Case "ClientFileList"
                    'dropped from client's list, so move the files
                    If Right(Data.Files(x), 1) = "\" Then
                        objFso.MoveFolder Left(Data.Files(x), Len(Data.Files(x)) - 1), ClientPath & itm.Text & "\"
                    Else
                        objFso.MoveFile Data.Files(x), ClientPath & itm.Text & "\" & objFso.GetFileName(Data.Files(x))
                    End If
                    RefreshClientView
            End Select
        Next x
    End If
    
End Sub

Private Sub ClientFileList_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    On Error Resume Next
    
    Dim itm As ListItem
    Set itm = ClientFileList.HitTest(x, y)
    If itm Is Nothing Then
        For x = 1 To ClientFileList.ListItems.Count
            ClientFileList.ListItems(x).Selected = False
            DoEvents
        Next x
        Exit Sub
    End If
    
    If ClientFileList.SelectedItem = itm Then Exit Sub

    For x = 1 To ClientFileList.ListItems.Count
        ClientFileList.ListItems(x).Selected = False
        DoEvents
    Next x
    
    If itm.Icon = "folder" Then
        itm.Selected = True
    Else
        
    End If

End Sub

Private Sub ClientFileList_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Call Data.SetData(, vbCFFiles)
    For x = 1 To ClientFileList.ListItems.Count
        If ClientFileList.ListItems(x).Selected = True Then
            If ClientFileList.ListItems(x).Icon = "folder" Then
                Data.Files.Add (ClientPath + ClientFileList.ListItems(x).Text & "\")
            Else
                Data.Files.Add (ClientPath + ClientFileList.ListItems(x).Text)
            End If
        End If
    Next x
    
    Source = "ClientFileList"
End Sub

Private Sub ClientFiles_GotFocus()
Call ShowFocus("Client")
End Sub














Private Sub Form_Activate()
If FirstRun = True Then
    'load client drive list
    Dim fs As New FileSystemObject, d As Drive, dc As Drives, itm As ComboItem
    Set dc = fs.Drives
    For Each d In dc
        S = d.DriveLetter
        Set itm = ClientDrives.ComboItems.Add(, , S & ":\", "drive")
        If d.DriveType = Fixed Then itm.Selected = True 'this will select the first fixed drive in the list
    Next
    
    lastfolder = ReadINI("Andromeda", "LastClientFolder", DLL())
    If lastfolder <> "" Then
        Set itm = ClientDrives.ComboItems.Add(, , lastfolder, "folder")
        itm.Selected = True
    End If
    
    Call ChowFromFolder(ClientFileList, ClientDrives.Text, "*.*")
    
    'Load server folders
    FirstRun = False
    Call frmMain.Winsock.SendData("SHAREDFOLDERS")
    
End If
End Sub

Sub StringToQue(theitems As String)
    'this code will loop through the path and make sure each parent
    'folder of the file exists. then, it will create any folders
    'that need to be created. (used for downloading folders)
    
    If theitems = "" Then Exit Sub
    If Not Mid(theitems, Len(theitems), 1) = "|" Then
    theitems = theitems & "|"
    End If
    
    Dim folders2create As New Collection
    
    For DoList = 1 To Len(theitems)
        thechars$ = thechars$ & Mid(theitems, DoList, 1)
        If Mid(theitems, DoList, 1) = "|" Then
            serverfile = Mid(thechars$, 1, Len(thechars$) - 1)
            clientfile = ClientPath + Right(serverfile, Len(serverfile) - Len(ClientPath))
            parentfld = objFso.GetParentFolderName(clientfile)
            Call Que.AddFile(serverfile & ">" & parentfld + "\")
            
            pf = objFso.GetParentFolderName(clientfile)
            Do While pf <> "": DoEvents
                If objFso.FolderExists(pf) = False Then
                    folders2create.Add pf
                End If
                pf = objFso.GetParentFolderName(pf)
            Loop
            thechars$ = ""
        End If
    Next DoList
    
    
    'Create folders (if needed)
    On Error Resume Next
    For x = folders2create.Count To 1 Step -1
        MkDir folders2create.item(x)
    Next x
    
    Set folders2create = Nothing
End Sub
Private Sub Form_Load()
    SetParent Me.hwnd, MDI()
    frmMain.Toolbar1.Buttons(3).Value = tbrPressed
    FirstRun = True
    
    Set Que = New FileQue
    Set UpQue = New FileQue
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    ClientFiles.Width = Me.Width / 2 - 100
    ServerFiles.Width = Me.Width / 2 - 100
    ClientDrives.Width = ClientFiles.Width - 300
    ServerDrives.Width = ServerFiles.Width - 300
    lblClientFiles.Width = ClientFiles.Width
    lblServerFiles.Width = ServerFiles.Width
    Toolbar1.Top = ClientDrives.Top + ClientDrives.Height + 25
    Toolbar1.Width = ClientFiles.Width
    Toolbar2.Top = ServerDrives.Top + ServerDrives.Height + 25
    Toolbar2.Width = ServerFiles.Width
    ClientFileList.Top = Toolbar1.Top + Toolbar1.Height + 25
    ClientFileList.Height = ClientFiles.Height - ClientFileList.Top - 100
    ClientFileList.Width = ClientFiles.Width - 75
    ServerFileList.Top = Toolbar2.Top + Toolbar2.Height + 25
    ServerFileList.Height = ServerFiles.Height - ServerFileList.Top - 100
    ServerFileList.Width = ServerFiles.Width - 75
    Status.Panels(1).Width = ClientFiles.Width - 20
    Status.Panels(3).Width = ServerFiles.Width
End Sub


















Private Sub Form_Unload(Cancel As Integer)
    frmMain.Toolbar1.Buttons(3).Value = tbrUnpressed
    Set Que = Nothing
    Set UpQue = Nothing
End Sub

Private Sub lblClientFiles_Click()
    Call ShowFocus("Client")
End Sub

Private Sub lblServerFiles_Click()
    Call ShowFocus("Server")
End Sub



Private Sub mnuClientMoveTo_Click()
    On Error Resume Next
    
    newpath = InputBox("Please enter the new path:", "Move")
    If Trim(newpath) = "" Then Exit Sub
    
    If Right(newpath, 1) <> "\" Then newpath = newpath + "\"
    
    If objFso.FolderExists(newpath) = False Then
        MsgBox "The specified path is invalid. Cannot move.", vbCritical, "Error"
        Exit Sub
    End If
    
    For x = 1 To ClientFileList.ListItems.Count
        If ClientFileList.ListItems(x).Selected = True Then
            If ClientFileList.ListItems(x).Icon = "folder" Then
                objFso.MoveFolder ClientPath + ClientFileList.ListItems(x).Text, newpath
            Else
                objFso.MoveFile ClientPath + ClientFileList.ListItems(x).Text, newpath + ClientFileList.ListItems(x).Text
            End If
        End If
    Next x
    
    RefreshClientView
    
End Sub

Private Sub mnuDelete_Click()
    If Focus = "Client" Then
        Call mnuDeleteClientFile_Click
    ElseIf Focus = "Server" Then
        Call mnuDeleteServerFile_Click
    End If
End Sub

Private Sub mnuDeleteClientFile_Click()
    On Error Resume Next
        If ClientFileList.SelectedItem Is Nothing Then Exit Sub
        
        For i = 1 To ClientFileList.ListItems.Count
            If ClientFileList.ListItems(i).Selected = True Then fcount = fcount + 1
        Next i
        If MsgBox("Delete " & fcount & " selected file(s) or folder(s)?" & vbCrLf, vbYesNo + vbQuestion, "Delete?") = vbYes Then
            Dim fs As New FileSystemObject
            Dim f1 As File
            Dim Fld As Folder
            For x = 1 To ClientFileList.ListItems.Count
                If ClientFileList.ListItems(x).Selected = True Then
                    If ClientFileList.ListItems(x).Icon = "folder" Then
                        Set Fld = fs.GetFolder(ClientPath + "\" + ClientFileList.ListItems(x) + "\")
                        Fld.Delete (True)
                    Else
                        Set f1 = fs.GetFile(ClientPath + "\" + ClientFileList.ListItems(x) + "\")
                        f1.Delete (True)
                    End If
                End If
            Next x
            RefreshClientView
        End If
End Sub

Private Sub mnuDeleteServerFile_Click()
    If ServerFileList.SelectedItem Is Nothing Then Exit Sub
    
    Dim str As String
    Dim fcount As Integer
    str = "DELETE="
    With ServerFileList
        For x = 1 To .ListItems.Count
            If .ListItems(x).Selected = True Then
                str = str & ServerPath & .ListItems(x).Text & "|"
                fcount = fcount + 1
            End If
        Next x
        retval = MsgBox("Delete " & fcount & " selected file(s) or folder(s)?" & vbCrLf, vbYesNo + vbQuestion, "Delete?")
        If retval = vbYes Then
            frmMain.Winsock.SendData (str)
        End If
    End With
    
    
End Sub

Private Sub mnuDownload_Click()
    
    For x = 1 To ServerFileList.ListItems.Count
        If ServerFileList.ListItems(x).Selected = True Then
            selnum = selnum + 1
            If ServerFileList.ListItems(x).Icon = "folder" Then
                Call GetServerFilesFromFolder(ServerPath + ServerFileList.ListItems(x).Text)
            Else
                Call Que.AddFile(ServerPath + ServerFileList.ListItems(x).Text & ">" & ClientPath)
            End If
        End If
    Next x
    
    If selnum = 0 Then Exit Sub
    
    If Que.IsDownloading = False Then
        Call Que.BeginDownloading
    End If
        
End Sub



Private Sub mnuExecuteServerProgram_Click()
    frmMain.Winsock.SendData "SPROCESS=" & ServerPath & ServerFileList.SelectedItem.Text
End Sub

Private Sub mnuNewClientFolder_Click()
    Dim NewFolder As String
    NewFolder = InputBox("Enter the new folder's name:", "New Folder", "NewFolder")
    If NewFolder = "" Then Exit Sub
    If IsValidFileName(NewFolder) = False Then
        MsgBox "The new folder name contains invalid characters.", vbCritical, "Invalid Folder Name"
        Exit Sub
    End If
    
    MkDir (ClientPath + NewFolder)
    
    RefreshClientView
    
End Sub

Private Sub mnuRenameClientFile_Click()
    Dim newname As String
    oldname = ClientFileList.SelectedItem.Text
    newname = InputBox("Enter the new name:", "Rename File: " & oldname, oldname)
    
    If Trim(newname) = "" Then Exit Sub
    If IsValidFileName(newname) = False Then
        MsgBox "New name contains illegal characters!", vbCritical, "Error"
        Exit Sub
    End If
    
    If ClientFileList.SelectedItem.Icon = "folder" Then
        Call RenameFolder(ClientPath + ClientFileList.SelectedItem.Text, newname)
    Else
        Call RenameFile(ClientPath + ClientFileList.SelectedItem.Text, newname)
    End If

    Call RefreshClientView
End Sub

Private Sub mnuServerMoveTo_Click()
    If ServerFileList.SelectedItem Is Nothing Then Exit Sub
    
    svrpath = InputBox("Please enter to directory to move this file to:", "Move File")
    If svrpath = "" Then Exit Sub
    
    If Right(svrpath, 1) <> "\" Then svrpath = svrpath + "\"
    
    For x = 1 To ServerFileList.ListItems.Count
        If ServerFileList.ListItems(x).Selected = True Then
            If ServerFileList.ListItems(x).Icon = "folder" Then
                frmMain.Winsock.SendData ("MOVE=" & ServerPath & ServerFileList.ListItems(x) & "\" & "|" & svrpath)
                WaitForMove = True
                While WaitForMove = True: DoEvents: Wend 'wait for server to move folder
            Else
                frmMain.Winsock.SendData ("MOVE=" & ServerPath & ServerFileList.ListItems(x) & "|" & svrpath)
                WaitForMove = True
                While WaitForMove = True: DoEvents: Wend 'wait for server to move file
            End If
        End If
    Next x
    
    Call RefreshServerView
End Sub

Private Sub mnuNewFolder_Click()
    Dim NewFolder As String
    NewFolder = InputBox("Enter the new folder's name:", "New Folder", "NewFolder")
    If NewFolder = "" Then Exit Sub
    If IsValidFileName(NewFolder) = False Then
        MsgBox "The new folder name contains invalid characters.", vbCritical, "Invalid Folder Name"
        Exit Sub
    End If
    
    Call frmMain.Winsock.SendData("NEWFOLDER=" & ServerPath & NewFolder)
    WaitForFolder = True
    While WaitForFolder = True: DoEvents: Wend
    RefreshServerView
End Sub

Private Sub mnuRenameServerFile_Click()
    Dim filename As String
    Dim new_fileName As String
    If ServerFileList.SelectedItem Is Nothing Then Exit Sub
    
    filename = ServerFileList.SelectedItem.Text
    
    new_fileName = InputBox("Enter the new file name.", "Rename file", filename)
    
    If Trim(new_fileName) = "" Then Exit Sub
    If IsValidFileName(new_fileName) = False Then
        MsgBox "New file name contains illegal characters!", vbCritical, "Error"
        Exit Sub
    End If
    
    Call frmMain.Winsock.SendData("RENAME=" & ServerPath & filename & "|" & new_fileName)
End Sub

Private Sub mnuUpload_Click()
    Dim item As String
    
    For x = 1 To ClientFileList.ListItems.Count
        If ClientFileList.ListItems(x).Selected = True Then
            selnum = selnum + 1
            item = ClientFileList.ListItems(x).Text
            
            If ClientFileList.ListItems(x).Icon = "folder" Then
                Call AddFolderToQue(ClientPath + item, ServerPath + item + "\", UpQue)
            Else
                UpQue.AddFile (ClientPath + item + ">" + ServerPath)
            End If
            
        End If
    Next x
    
    If selnum = 0 Then Exit Sub
    
    If UpQue.IsUploading = False Then
        UpQue.BeginUploading
    End If
End Sub

Private Sub mnuDetails_Click()
    ClientFileList.View = lvwReport
    ServerFileList.View = lvwReport
    
    mnuSmallIcons.Checked = False
    mnuLargeIcons.Checked = False
    mnuList.Checked = False
    mnuDetails.Checked = True
End Sub

Private Sub mnuLargeIcons_Click()
    ClientFileList.View = lvwIcon
    ServerFileList.View = lvwIcon
    
    mnuSmallIcons.Checked = False
    mnuLargeIcons.Checked = True
    mnuList.Checked = False
    mnuDetails.Checked = False
End Sub


Private Sub mnuList_Click()
    ClientFileList.View = lvwList
    ServerFileList.View = lvwList
    
    mnuSmallIcons.Checked = False
    mnuLargeIcons.Checked = False
    mnuList.Checked = True
    mnuDetails.Checked = False
End Sub



Private Sub mnuViewRefresh_Click()
    Call ChowFromFolder(ClientFileList, ClientPath, "*.*")
    Call frmMain.Winsock.SendData("DIR " + ServerPath)
    ServerFileList.MousePointer = 13
End Sub

Private Sub mnuSmallIcons_Click()
    ClientFileList.View = lvwSmallIcon
    ServerFileList.View = lvwSmallIcon
    
    mnuSmallIcons.Checked = True
    mnuLargeIcons.Checked = False
    mnuList.Checked = False
    mnuDetails.Checked = False
End Sub





Private Sub Que_DownloadsComplete(NumFiles As Integer)
    If frmDownload.CloseWhenDone.Value = 1 Then
        Unload frmDownload
    End If
    RefreshClientView
End Sub

Private Sub ServerDrives_Click()
    Path = ServerDrives.SelectedItem.Text
    If Path = "" Then Exit Sub
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    
    Call frmMain.Winsock.SendData("DIR " + Path)
    ServerFileList.MousePointer = 13
    ServerFileList.SetFocus
End Sub

Private Sub ServerDrives_GotFocus()
ShowFocus "Server"
End Sub





Private Sub ServerDrives_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Path = ServerDrives.Text
        If Path = "" Then Exit Sub
        If Right(Path, 1) <> "\" Then Path = Path + "\"
        
        Call frmMain.Winsock.SendData("DIR " + Path)
        ServerFileList.MousePointer = 13
        ServerFileList.SetFocus
    End If
End Sub

Private Sub ServerFileList_AfterLabelEdit(Cancel As Integer, NewString As String)
    If IsValidFileName(NewString) = False Then
        MsgBox "New file name contains illegal characters!", vbCritical, "Error"
        Cancel = 1
        Exit Sub
    End If
    
    Call frmMain.Winsock.SendData("RENAME=" & ServerPath & ServerFileList.SelectedItem.Text & "|" & Trim(NewString))
End Sub

Private Sub ServerFileList_DblClick()
    If ServerFileList.SelectedItem Is Nothing Then Exit Sub
    If ServerFileList.SelectedItem.Icon = "folder" Then
        frmMain.Winsock.SendData ("DIR " + ServerPath + ServerFileList.SelectedItem.Text + "\")
skipit:
    Else
        Call mnuDownload_Click
    End If
    
    
End Sub



Private Sub ServerFileList_GotFocus()
ShowFocus "Server"
End Sub






Private Sub ServerFileList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set ServerFileList.SelectedItem = Nothing
End Sub

Private Sub ServerFileList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        For x = 1 To ServerFileList.ListItems.Count
            If ServerFileList.ListItems(x).Selected = True Then n = n + 1
        Next x
        If n = 0 Then Exit Sub
        
        If n > 1 Then
            mnuRenameServerFile.Enabled = False
        Else
            mnuRenameServerFile.Enabled = True
        End If
        
        On Error Resume Next
        If objFso.GetExtensionName(ServerFileList.SelectedItem.Text) = "exe" Then
            mnuExecuteServerProgram.Enabled = True
        Else
            mnuExecuteServerProgram.Enabled = False
        End If
        
        Call Me.PopupMenu(mnuServer)
    End If
End Sub



Private Sub ServerFileList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itm As ListItem
    Set itm = ServerFileList.HitTest(x, y)
    If itm Is Nothing Then Exit Sub
    If itm.Icon = "folder" Then
        'item(s) dropped in folder
            Select Case Source
                Case "ServerFileList"
                    'dropped from server's list, so send the "MOVE" command
                    For x = 1 To Data.Files.Count
                        frmMain.Winsock.SendData ("MOVE=" & ServerPath & Data.Files(x) & "|" & ServerPath & itm.Text & "\") ' & Data.Files(x))
                        WaitForMove = True
                        While WaitForMove = True: DoEvents: Wend 'wait for server to move file
                    Next x
                    Call RefreshServerView
                    
            End Select
        
    End If
End Sub

Private Sub ServerFileList_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    On Error Resume Next
    Dim itm As ListItem
    Set itm = ServerFileList.HitTest(x, y)
    
    If itm Is Nothing Then
        For x = 1 To ServerFileList.ListItems.Count
            ServerFileList.ListItems(x).Selected = False
            DoEvents
        Next x
        Exit Sub
    End If
    
    If ServerFileList.SelectedItem = itm Then Exit Sub

    For x = 1 To ServerFileList.ListItems.Count
        ServerFileList.ListItems(x).Selected = False
        DoEvents
    Next x
    
    If itm.Icon = "folder" Then
        itm.Selected = True
    Else
        
    End If
End Sub


Private Sub ServerFileList_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Call Data.SetData(, vbCFFiles)
    For x = 1 To ServerFileList.ListItems.Count
        If ServerFileList.ListItems(x).Selected = True Then
            If ServerFileList.ListItems(x).Icon = "folder" Then
                Data.Files.Add (ServerFileList.ListItems(x) & "\")
            Else
                Data.Files.Add (ServerFileList.ListItems(x).Text)
            End If
        End If
    Next x
    
    Source = "ServerFileList"
End Sub

Private Sub ServerFiles_Click()
Call ShowFocus("Server")
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Call ShowFocus("Client")
Select Case Button.index
    Case 1
        'Go Up
        Me.MousePointer = 13
        Set f = New FileSystemObject
        p$ = f.GetParentFolderName(ClientPath)
        If p$ = "" Then Exit Sub
        Call ChowFromFolder(ClientFileList, p$ + "\", "*.*")
        Me.MousePointer = 0
    Case 3
        'Delete
        Call mnuDeleteClientFile_Click
    Case 4
        'Upload
        Call mnuUpload_Click
    Case 6
        'Large Icons
        Call mnuLargeIcons_Click
    Case 7
        'Small Icons
        Call mnuSmallIcons_Click
    Case 8
        'List
        Call mnuList_Click
    Case 9
        'Details
        Call mnuDetails_Click
    Case 11
        'New Folder
        Call mnuNewClientFolder_Click
End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Call ShowFocus("Server")
Select Case Button.index
    Case 1
        'Go Up
        Me.MousePointer = 13
        Set f = New FileSystemObject
        p$ = f.GetParentFolderName(ServerPath)
        If p$ = "" Then Exit Sub
        If Right(p$, 1) <> "\" Then p$ = p$ + "\"
        Call frmMain.Winsock.SendData("DIR " + p$)
        
skip:
        Me.MousePointer = 0
        ServerFileList.MousePointer = 0
    Case 3
        'Delete files
        Call mnuDeleteServerFile_Click
    Case 4
        'Upload
        Call mnuUpload_Click
    Case 6
        'Large Icons
        Call mnuLargeIcons_Click
    Case 7
        'Small Icons
        Call mnuSmallIcons_Click
    Case 8
        'List
        Call mnuList_Click
    Case 9
        'Details
        Call mnuDetails_Click
    Case 11
        'New Folder
        Call mnuNewFolder_Click
End Select
End Sub


Private Sub UpQue_UploadsComplete(NumFiles As Integer)
    If frmUpload.CloseWhenDone.Value = 1 Then
        Unload frmUpload
    End If
    RefreshServerView
End Sub


