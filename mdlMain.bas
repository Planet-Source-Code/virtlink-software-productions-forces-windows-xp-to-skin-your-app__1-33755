Attribute VB_Name = "mdlMain"
Public Declare Sub InitCommonControls Lib "comctl32" ()

Public strProject As String

Public intModules As Integer
Public strModulePath(1000) As String
Public strModuleName(1000) As String

Public intForms As Integer
Public strFormPath(1000) As String
Public Sub GetModules()
    Dim strA As String
    
    intModules = 0
    
    f = FreeFile
    Open strProject For Input As #f
        While Not EOF(f)
            Line Input #f, strA
            If LCase(Left(strA, 6)) = "module" Then
                intModules = intModules + 1
                strA = Adapt(Right(strA, Len(strA) - InStrRev(strA, "=")))
                strModulePath(intModules) = Right(strA, Len(strA) - InStr(strA, ";"))
                strModuleName(intModules) = Left(strA, InStr(strA, ";") - 1)
            End If
        Wend
    Close
End Sub

Public Sub GetForms()
    Dim strA As String
    
    intForms = 0
    
    f = FreeFile
    Open strProject For Input As #f
        While Not EOF(f)
            Line Input #f, strA
            If LCase(Left(strA, 4)) = "form" Then
                intForms = intForms + 1
                strFormPath(intForms) = Adapt(Right(strA, Len(strA) - InStrRev(strA, "=")))
            End If
        Wend
    Close
End Sub

Public Sub DoForms()
    Dim q As Integer
    
    For q = 1 To intForms
        DoForm Left(strProject, InStrRev(strProject, "\")) & Trim(strFormPath(q))
    Next
End Sub

Public Sub DoModule(strPath As String)
    Dim f As Integer, strA As String, strB As String
    
    f = FreeFile
    Open strPath For Input As #f
        Line Input #f, strA
        strA = strA & vbCrLf & "Public Declare Sub InitCommonControls Lib " & Chr(34) & "comctl32" & Chr(34) & " ()" & vbCrLf
        While Not EOF(f)
            Line Input #f, strB
            strA = strA & vbCrLf & strB
        Wend
    Close
    
    FileCopy strPath, strPath & ".bck"
    
    f = FreeFile
    Open strPath For Output As #f
        Print #f, strA
    Close
End Sub
Public Sub DoForm(strPath As String)
    Dim f As Integer, strA As String, strB As String, b As Boolean
    
    f = FreeFile
    Open strPath For Input As #f
        While Not EOF(f)
            Line Input #f, strB
            strA = strA & vbCrLf & strB
            If strB = "Private Sub Form_Initialize()" And b = False Then
                strA = strA & vbCrLf & "    InitCommonControls" & vbCrLf
                b = True
            End If
        Wend
        If b = False Then
            strA = strA & vbCrLf & "Private Sub Form_Initialize()" & vbCrLf & "    InitCommonControls" & vbCrLf & "End Sub" & vbCrLf
        End If
    Close
    
    FileCopy strPath, strPath & ".bck"
    
    f = FreeFile
    Open strPath For Output As #f
        Print #f, strA
    Close
End Sub
Public Function GetExeName() As String
    Dim strA As String
    
    f = FreeFile
    Open strProject For Input As #f
        While Not EOF(f)
            Line Input #f, strA
            If LCase(Left(strA, 9)) = "exename32" Then
                GetExeName = Adapt(Right(strA, Len(strA) - InStrRev(strA, "=")))
                Exit Function
            End If
        Wend
    Close
End Function

Public Function Adapt(strAdapt As String) As String
    If Left(strAdapt, 1) = Chr(34) Then strAdapt = Right(strAdapt, Len(strAdapt) - 1)
    If Right(strAdapt, 1) = Chr(34) Then strAdapt = Left(strAdapt, Len(strAdapt) - 1)
    Adapt = strAdapt
End Function

Public Sub MakeManifest(strPath As String)
    Dim f As Integer, a As String, strA As String, strB As String, b As Long
    
    f = FreeFile
    Open App.Path & "\projupdt.exe.manifest" For Input As #f
        While Not EOF(f)
            Line Input #f, strB
            strA = strA & vbCrLf & strB
        Wend
    Close
    
    a = GetExeName
    If a = "" Then a = Right(strProject, Len(strProject) - InStrRev(strProject, "\"))
        'a = Left(strProject, InStrRev(strProject, ".") - 1)
    'End If
    a = LCase(Left(a, InStrRev(a, ".") - 1))
    
    b = InStr(1, LCase(strA), "projupdt")
    strA = Left(strA, b - 1) & a & Right(strA, Len(strA) - b - 8 + 1)
    
    f = FreeFile
    Open strPath & a & ".exe.manifest" For Output As #f
        Print #f, strA
    Close
End Sub
