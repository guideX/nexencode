'nexENCODE Studio 5.0 Alpha 1.3
'October 6th, 2013
Option Explicit On
Option Strict On
Public Class clsScripting
    Enum eVariableTypes
        vString = 1
        vInteger = 2
    End Enum

    Structure gVariable
        Public vScope As String
        Public vName As String
        Public vString As String
        Public vInteger As Integer
        Public vVariableType As eVariableTypes
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
                        .vName = _Variable.vName
                        .vScope = _Variable.vScope
                        .vString = _Variable.vString
                        .vInteger = _Variable.vInteger
                        b = True
                    End If
                End With
            Next i
            If b = False Then
                lVariables.vCount = lVariables.vCount + 1
                ReDim Preserve lVariables.vVariable(lVariables.vCount)
                With lVariables.vVariable(lVariables.vCount)
                    .vName = _Variable.vName
                    .vScope = _Variable.vScope
                    .vString = _Variable.vString
                    .vInteger = _Variable.vInteger
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
            Dim splt2() As String, lVariable As New gVariable
            For Each lLine As String In splt
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