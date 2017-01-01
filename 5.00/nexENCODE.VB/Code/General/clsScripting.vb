'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Public Class clsScripting
    Enum eVariableTypes
        vNothing = 0
        vString = 1
        vInteger = 2
        vForm = 3
        vButton = 4
    End Enum

    Structure gVariable
        Public vButton As Button
        Public vScope As String
        Public vName As String
        Public vString As String
        Public vInteger As Integer
        Public vForm As frmScriptedForm
        Public vVariableType As eVariableTypes
        Public vParentForm As String
    End Structure

    Structure gVariables
        Public vCount As Integer
        Public vVariable() As gVariable
    End Structure

    Public Event ProcessError(lError As String, lSub As String)
    Private WithEvents lFiles As New clsFiles
    Private lCodeFile As String
    Private lVariables As gVariables

    Public Function DoesLineMatch(lLine As String, lCompare As String) As Boolean
        Try
            If LCase(Trim(lLine)) = LCase(Trim(lCompare)) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function DoesLineMatch(lLine As String, lCompare As String) As Boolean")
            Return Nothing
        End Try
    End Function

    Public Sub AddUpdateVariable(_Variable As gVariable)
        Try
            Dim b As Boolean = False
            For i As Integer = 1 To lVariables.vCount
                With lVariables.vVariable(i)
                    If LCase(Trim(.vName)) = LCase(Trim(_Variable.vName)) And LCase(Trim(.vScope)) = LCase(Trim(_Variable.vScope)) Then
                        If Not (_Variable.vParentForm Is Nothing) Then
                            If Not (_Variable.vParentForm.Length = 0) Then .vParentForm = _Variable.vParentForm
                        End If
                        If Not (_Variable.vButton Is Nothing) Then .vButton = _Variable.vButton
                        If Not (String.IsNullOrEmpty(_Variable.vName)) Then .vName = _Variable.vName
                        If Not (String.IsNullOrEmpty(_Variable.vScope)) Then .vScope = _Variable.vScope
                        If Not (String.IsNullOrEmpty(_Variable.vString)) Then .vString = _Variable.vString
                        If Not _Variable.vInteger = 0 Then .vInteger = _Variable.vInteger
                        If Not (_Variable.vForm Is Nothing) Then .vForm = _Variable.vForm
                        If _Variable.vVariableType <> eVariableTypes.vNothing Then .vVariableType = _Variable.vVariableType
                        b = True
                    End If
                End With
            Next i
            If b = False Then
                lVariables.vCount = lVariables.vCount + 1
                ReDim Preserve lVariables.vVariable(lVariables.vCount)
                With lVariables.vVariable(lVariables.vCount)
                    If Not (_Variable.vParentForm Is Nothing) Then
                        If Not (_Variable.vParentForm.Length = 0) Then .vParentForm = _Variable.vParentForm
                    End If
                    If Not (_Variable.vButton Is Nothing) Then .vButton = _Variable.vButton
                    If Not (String.IsNullOrEmpty(_Variable.vName)) Then .vName = _Variable.vName
                    If Not (String.IsNullOrEmpty(_Variable.vScope)) Then .vScope = _Variable.vScope
                    If Not (String.IsNullOrEmpty(_Variable.vString)) Then .vString = _Variable.vString
                    If _Variable.vInteger <> 0 Then .vInteger = _Variable.vInteger
                    If Not (_Variable.vForm Is Nothing) Then .vForm = _Variable.vForm
                    If _Variable.vVariableType <> eVariableTypes.vNothing Then .vVariableType = _Variable.vVariableType
                End With
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub AddUpdateVariable(lScope As String, lName As String, lValue As String)")
        End Try
    End Sub

    Public Function ProcessReplaceVariables(lLine As String) As String
        Try
            For i As Integer = 1 To lVariables.vCount
                lLine = Replace(lLine, lVariables.vVariable(i).vName, lVariables.vVariable(i).vString)
            Next i
            Return lLine
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Function ProcessReplaceVariables(lLine As String) As String")
            Return Nothing
        End Try
    End Function

    Public Sub ProcessCode(lCode As String, lScope As String)
        Try
            Dim splt() As String = Split(lCode, vbCrLf), msg As String, msg2 As String = ""
            Dim splt2() As String, splt3() As String, lVariable As New gVariable
            For Each lLine As String In splt
                If (Not (Left(lLine, 2) = "//")) Then
                    If Left(lLine.ToLower(), 7) = "action " Then
                        msg = Right(lLine, lLine.Length - 7)
                        splt3 = Split(msg, ",")
                        For i As Integer = 1 To lVariables.vVariable.Count
                            If Not (String.IsNullOrEmpty(lVariables.vVariable(i).vName)) Then
                                If (splt3(0).Trim() = lVariables.vVariable(i).vName.Trim()) Then
                                    Select Case lVariables.vVariable(i).vVariableType
                                        Case eVariableTypes.vForm
                                            Select Case splt3(1).Replace(Chr(34), "").Trim().ToLower()
                                                Case "show"
                                                    lVariables.vVariable(i).vForm.Show()
                                                Case "hide"
                                                    lVariables.vVariable(i).vForm.Hide()
                                            End Select
                                    End Select
                                    Exit For
                                End If
                            End If
                        Next i
                    End If
                    If Left(lLine.ToLower, 4) = "set " Then
                        msg = Right(lLine, lLine.Length - 4)
                        splt3 = Split(msg, ",")
                        If (UBound(splt3)) = 2 Then
                            For i As Integer = 1 To lVariables.vVariable.Count
                                If (lVariables.vVariable(i).vVariableType <> eVariableTypes.vNothing) Then
                                    If (lVariables.vVariable(i).vName.ToLower() = splt3(0).Trim().ToLower()) Then
                                        If (Left(splt3(2).Trim(), 1) = Chr(34) And Right(splt3(2).Trim(), 1) = Chr(34)) Then
                                            Select Case lVariables.vVariable(i).vVariableType
                                                Case eVariableTypes.vButton
                                                    Select Case splt3(1).Trim().ToLower().Replace(Chr(34), "")
                                                        Case "name"
                                                            MsgBox(splt3(1))
                                                    End Select
                                                Case eVariableTypes.vForm
                                                    Select Case splt3(1).Trim().ToLower().Replace(Chr(34), "")
                                                        Case "width"
                                                            If (IsNumeric(splt3(2).Replace(Chr(34), ""))) Then lVariables.vVariable(i).vForm.Width = CType(splt3(2).Replace(Chr(34), ""), Integer)
                                                        Case "height"
                                                            If (IsNumeric(splt3(2).Replace(Chr(34), ""))) Then lVariables.vVariable(i).vForm.Height = CType(splt3(2).Replace(Chr(34), ""), Integer)
                                                        Case "icon"
                                                            Dim bmp As System.Drawing.Image = Bitmap.FromFile(splt3(2).Replace("$apppath", Application.StartupPath & "\").Replace(Chr(34), ""))
                                                            Dim thumb As System.Drawing.Image = bmp.GetThumbnailImage(64, 64, Nothing, IntPtr.Zero)
                                                            Dim icn = Icon.FromHandle(CType(thumb, Bitmap).GetHicon())
                                                            lVariables.vVariable(i).vForm.Icon = icn
                                                        Case "name"
                                                            lVariables.vVariable(i).vForm.FormName = splt3(2).Replace(Chr(34), "").Trim()
                                                        Case "title"
                                                            lVariables.vVariable(i).vForm.FormTitle = splt3(2).Replace(Chr(34), "").Trim()
                                                    End Select
                                            End Select
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next i
                        End If
                    End If
                    If Left(LCase(lLine), 4) = "btn " Then
                        splt2 = Split(lLine, " = ")
                        lVariable.vName = Replace(splt2(0), "btn ", "")
                        lVariable.vButton = New Button()
                        lVariable.vScope = lScope
                        lVariable.vVariableType = eVariableTypes.vButton
                        AddUpdateVariable(lVariable)
                    End If
                    If Left(LCase(lLine), 4) = "frm " Then
                        splt2 = Split(lLine, " = ")
                        lVariable.vName = Replace(splt2(0), "frm ", "")
                        lVariable.vForm = New frmScriptedForm()
                        lVariable.vScope = lScope
                        lVariable.vVariableType = eVariableTypes.vForm
                        AddUpdateVariable(lVariable)
                    End If
                    If Left(LCase(lLine), 4) = "int " Then
                        splt2 = Split(lLine, " = ")
                        lVariable.vName = Replace(splt2(0), "int ", "")
                        lVariable.vInteger = CInt(Trim(splt2(1)))
                        lVariable.vScope = lScope
                        lVariable.vVariableType = eVariableTypes.vInteger
                        AddUpdateVariable(lVariable)
                    End If
                    If Left(LCase(lLine), 4) = "var " Then
                        splt2 = Split(lLine, " = ")
                        lVariable.vName = Replace(splt2(0), "var ", "")
                        lVariable.vString = Replace(splt2(1), Chr(34), "")
                        lVariable.vScope = lScope
                        lVariable.vVariableType = eVariableTypes.vString
                        AddUpdateVariable(lVariable)
                    End If
                    If Left(LCase(lLine), 6) = "exit()" Then
                        Application.Exit()
                    End If
                    If Left(LCase(lLine), 10) = "minimize()" Then
                        frmMain.WindowState = FormWindowState.Minimized
                    End If
                    If Left(LCase(lLine), 10) = "maximize()" Then
                        MsgBox("TODO")
                    End If
                    If Left(LCase(lLine), 13) = "createobject()" Then
                        msg = Right(lLine, Len(lLine) - 14)
                        MsgBox(msg)
                    End If
                    If Left(LCase(lLine), 6) = "msgbox" Then
                        msg = Right(lLine, Len(lLine) - 6)
                        If InStr(msg, "(") <> 0 And InStr(msg, ")") <> 0 Then
                            If Left(msg, 1) = "(" Then
                                msg = Right(msg, Len(msg) - 1)
                                If Left(msg, 1) = Chr(34) Then
                                    msg = Right(msg, Len(msg) - 1)
                                    For Each lChar As Char In msg
                                        If lChar <> Chr(34) And lChar <> ")" Then
                                            msg2 = msg2 & lChar
                                        End If
                                    Next lChar
                                    msg2 = ProcessReplaceVariables(msg2)
                                    MsgBox(msg2, MsgBoxStyle.Information)
                                End If
                            End If
                        End If
                    End If
                End If
            Next lLine
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ProcessCode(lCode As String)")
        End Try
    End Sub

    Public Sub ProcessPrimitive(lPrimitive As String)
        Try
            Dim msg As String, msg2 As String = "", b As Boolean
            If lFiles.DoesFileExist(lCodeFile) = True Then
                msg = System.IO.File.ReadAllText(lCodeFile)
                For Each lLine As String In Split(msg, vbCrLf)
                    If b = False Then
                        If DoesLineMatch("Primitive " & lPrimitive, lLine) = True Then
                            b = True
                        End If
                    Else
                        If DoesLineMatch("End " & lPrimitive, lLine) = True Then
                            b = False
                            Exit For
                        Else
                            If Len(msg2) <> 0 Then
                                msg2 = msg2 & vbCrLf & lLine
                            Else
                                msg2 = lLine
                            End If
                        End If
                    End If
                Next lLine
                ProcessCode(msg2, lPrimitive)
            End If
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub ProcessPrimitive(lPrimitive As String, lCodeFile As String)")
        End Try
    End Sub

    Public Sub New(_CodeFile As String)
        Try
            lCodeFile = _CodeFile
        Catch ex As Exception
            RaiseEvent ProcessError(ex.Message, "Public Sub New(_CodeFile As String)")
        End Try
    End Sub
End Class