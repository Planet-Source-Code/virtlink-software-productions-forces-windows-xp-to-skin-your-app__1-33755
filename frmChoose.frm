VERSION 5.00
Begin VB.Form frmChoose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose module"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmChoose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox lstModules 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Caption         =   "VirtLink Project Updater has to make some changes to a module. Which module do you want to use?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    DoModule Left(strProject, InStrRev(strProject, "\")) & Trim(strModulePath(lstModules.ListIndex + 1))
    Unload Me
    DoForms
    MakeManifest Left(strProject, InStrRev(strProject, "\"))
    MsgBox "Updating project successfully completed!", vbInformation, "VirtLink Project Updater"
    frmMain.Enabled = True
End Sub


Private Sub Form_Load()
    Dim q As Integer
    
    For q = 1 To intModules
        lstModules.AddItem strModuleName(q)
    Next
End Sub


Private Sub Form_Initialize()
    InitCommonControls

    InitCommonControls

    InitCommonControls

    InitCommonControls
End Sub

