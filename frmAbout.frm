VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Info"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblRights 
      Alignment       =   2  'Center
      Caption         =   "All rights reserved."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "The icon of Project Updater is property of Microsoft Corp."
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblBuild 
      Alignment       =   2  'Center
      Caption         =   "Build: 1234"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      Caption         =   "Auteur: Daniël Pelsmaeker"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2002 - VirtLink Software Productions"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "VirtLink SavePort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strKey As String
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls

    InitCommonControls

    InitCommonControls

    'InitCommonControls
End Sub

Private Sub Form_Load()
    lblBuild.Caption = "Build: " & App.Revision
    lblLabels(0).Caption = "VirtLink Project Updater v. " & App.Major & "." & App.Minor
End Sub


