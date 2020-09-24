VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VirtLink Project Updater"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Project"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtProject 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub


Private Sub cmdBrowse_Click()
    On Error GoTo 0
    On Error GoTo stpError
    
    dlgCommon.CancelError = True
    dlgCommon.DefaultExt = "vbp"
    dlgCommon.DialogTitle = "Browse for Visual Basic Project"
    dlgCommon.Filter = "Visual basic Project (*.vbp)|*.vbp"
    dlgCommon.Flags = &H1000 And &H4
    dlgCommon.ShowOpen
    
    txtProject.Text = dlgCommon.FileTitle
    strProject = dlgCommon.FileName
    
    Exit Sub
stpError:
    Select Case Err
        Case 32755
    End Select
End Sub


Private Sub cmdUpdate_Click()
    Dim f As Integer
    Me.Enabled = False
    GetModules
    GetForms
    If intModules = 0 Then
        ans = MsgBox("No modules do exist in your project. Do you want to create one?", vbYesNoCancel + vbQuestion, "VirtLink Project Updater")
        If ans = vbYes Then
            strModulePath(1) = Left(strProject, InStrRev(strProject, "\")) & "mMain.bas"
            strModuleName(1) = "mMain"
            f = FreeFile
            Open strModulePath(1) For Output As #f
                Print #f, "Attribute VB_Name = " & Chr(34) & "mMain" & Chr(34)
            Close
            f = FreeFile
            Open strProject For Input As #f
                strA = "Module=mMain; mMain.bas" & vbCrLf
                While Not EOF(f)
                    Line Input #f, strB
                    strA = strA & strB & vbCrLf
                Wend
            Close #f
            
            FileCopy strProject, strProject & ".bck"
            
            f = FreeFile
            Open strProject For Output As #f
                Print #f, strA
            Close #f
            DoModule strModulePath(1)
            DoForms
            MakeManifest Left(strProject, InStrRev(strProject, "\"))
            MsgBox "Updating project successfully completed!", vbInformation, "VirtLink Project Updater"
            Me.Enabled = True
        Else
            Me.Enabled = True
            Exit Sub
        End If
    ElseIf intForms = 0 Then
        MsgBox "You don't have any forms in your project.", vbExclamation, "VirtLink Project Updater"
        Me.Enabled = True
        Exit Sub
    ElseIf intModules > 1 Then
        Load frmChoose
        frmChoose.Show
    Else
        DoModule Left(strProject, InStrRev(strProject, "\")) & Trim(strModulePath(1))
        DoForms
        MakeManifest Left(strProject, InStrRev(strProject, "\"))
        MsgBox "Updating project successfully completed!", vbInformation, "VirtLink Project Updater"
        Me.Enabled = True
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

