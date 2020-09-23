VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "bout Andromeda RFS"
   ClientHeight    =   3120
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
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClose 
      Caption         =   "Great"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   530
      Left            =   80
      Picture         =   "frmAbout.frx":014A
      ScaleHeight     =   465
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label5 
      Caption         =   "Web Site: www.induhviduals.com/andromeda"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Andrew: d_lederman@europe.com"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Ryan: rlederman@mad.scientist.com"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Written by Andrew and Ryan Lederman"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":1FBC
      Height          =   1215
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
Me.Hide
End Sub


