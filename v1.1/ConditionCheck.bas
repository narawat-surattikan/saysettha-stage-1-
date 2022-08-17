Attribute VB_Name = "ConditionCheck"
'' -1 : unset  0 : true   1 : false
Public Function Check(ByVal strInput As String) As Integer
    Dim fixed As String
    Dim stackCond(1 To 1000000, 1 To 3) As Variant
    Dim BooleanVal() As Variant
    Dim arrayCount, i As Variant

    strInput = StringRefine.ReplaceStr(strInput, "<=", "`")
    strInput = StringRefine.ReplaceStr(strInput, ">=", "~")
    strInput = StringRefine.ReplaceStr(strInput, "!=", "!")
    strInput = StringRefine.ReplaceStr(strInput, "<>", "!")
    strInput = StringRefine.ReplaceStr(strInput, "&&", "&")
    strInput = StringRefine.ReplaceStr(strInput, "||", "|")
    strInput = StringRefine.ReplaceStr(strInput, "and", "&")
    strInput = StringRefine.ReplaceStr(strInput, "or", "|")
    strInput = StringRefine.ReplaceStr(strInput, "==", "=")
    
    strInput = Module1.RefineConditionString(strInput)
        
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = "|{" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
                strInput = Ins(strInput, "| ", i)
            End If
        End If
    Next i
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = "&{" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
                strInput = Ins(strInput, "& ", i)
            End If
        End If
    Next i
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = "{#" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
                strInput = Ins(strInput, "{ ", i)
            End If
        End If
    Next i
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 1)) = "}" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
                strInput = Ins(strInput, " }", i)
            End If
        End If
    Next i
    strInput = TestRPNForCondition.Calc(strInput)
    
    
    Dim k() As String
    k = DoSplitObjectNew(strInput, " ")
    ''Slide1.console.Text = ""
    For i = 1 To CInt(k(0))
        
        fixed = StringRefine.Trim(k(i))
        If fixed <> "" Then
            If StringRefine.CustomRegexChecker(fixed, "^[$].*$") = 0 Then
                If Recognise.RVal(fixed)(1) = "-1" Then
                    GoTo HandleError
                End If
                arrayCount = arrayCount + 1
                stackCond(arrayCount, 1) = "-1"
                stackCond(arrayCount, 2) = Recognise.RVal(fixed)(2)
            ElseIf Mid(fixed, 1, 1) = Chr(34) Then
                If Recognise.RString(fixed)(1) = "-1" Then
                    GoTo HandleError
                End If
                arrayCount = arrayCount + 1
                stackCond(arrayCount, 1) = "-1"
                stackCond(arrayCount, 2) = Recognise.RString(fixed)(2)
            ElseIf StringRefine.VerifyNumber(fixed) = 0 Then
                arrayCount = arrayCount + 1
                stackCond(arrayCount, 1) = "-1"
                stackCond(arrayCount, 2) = CStr(fixed)
                stackCond(arrayCount, 3) = "int"
            ElseIf StringRefine.CustomRegexChecker(LCase(fixed), "^calc[ ]*[(].+[)][ ]*$") = 0 Then
                refineName = fixed
                refineName = Right(refineName, Len(refineName) - 4)
                refineName = StringRefine.Trim(refineName)
                refineName = Right(refineName, Len(refineName) - 1)
                refineName = Left(refineName, Len(refineName) - 1)
                tempvalue = Calculator.Calc(refineName)
                If tempvalue <> "Math Error" Then
                    arrayCount = arrayCount + 1
                    stackCond(arrayCount, 1) = "-1"
                    stackCond(arrayCount, 2) = CStr(tempvalue)
                    stackCond(arrayCount, 3) = "int"
                Else
                    GoTo HandleError
                End If
            ElseIf StringRefine.CustomRegexChecker(LCase(fixed), "^concat[ ]*[(].+[)][ ]*$") = 0 Then
                tempvalue = ConcatString(fixed)(1)
                If tempvalue <> "e" Then
                    arrayCount = arrayCount + 1
                    stackCond(arrayCount, 1) = "-1"
                    stackCond(arrayCount, 2) = CStr(ConcatString(fixed)(2))
                Else
                    GoTo HandleError
                End If
            ElseIf fixed = "=" Or fixed = "<" Or fixed = ">" Or fixed = "!" Or fixed = "`" Or fixed = "~" Then
                 Select Case fixed
                    Case "="
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) = CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) = stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                    Case "~"
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) >= CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) >= stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                    Case "`"
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) <= CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) <= stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                    Case ">"
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) > CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) > stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                    Case "<"
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) < CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) < stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                    Case "!"
                        If _
                            ( _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount - 1, 2))) = 0 And _
                                StringRefine.VerifyNumber(CStr(stackCond(arrayCount, 2))) = 0 And _
                                (CStr(stackCond(arrayCount, 3)) = "int" Or CStr(stackCond(arrayCount - 1, 3)) = "int") _
                            ) _
                        Then
                            If CSng(stackCond(arrayCount - 1, 2)) <> CSng(stackCond(arrayCount, 2)) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        Else
                            If stackCond(arrayCount - 1, 2) <> stackCond(arrayCount, 2) Then
                                GoTo AssignTrue
                            Else
                                GoTo AssignFalse
                            End If
                        End If
                 End Select
            ElseIf fixed = "&" Then
                If stackCond(arrayCount - 1, 1) = "0" And stackCond(arrayCount, 1) = "0" Then
                    GoTo AssignTrue
                Else
                    GoTo AssignFalse
                End If
            ElseIf fixed = "|" Then
                If stackCond(arrayCount - 1, 1) = "0" Or stackCond(arrayCount, 1) = "0" Then
                    GoTo AssignTrue
                Else
                    GoTo AssignFalse
                End If
            Else
                GoTo HandleError
            End If
        End If
BackToForLoop:
    Next
    Check = stackCond(1, 1)
    Exit Function
HandleError:
MsgBox "Oh no your computer has virus"
Check = 2
Exit Function
AssignTrue:
stackCond(arrayCount - 1, 1) = "0"
stackCond(arrayCount, 1) = "-1"
stackCond(arrayCount, 2) = ""
arrayCount = arrayCount - 1
GoTo BackToForLoop
AssignFalse:
stackCond(arrayCount - 1, 1) = "1"
stackCond(arrayCount, 1) = "-1"
stackCond(arrayCount, 2) = ""
arrayCount = arrayCount - 1
GoTo BackToForLoop
End Function
Public Sub j()
Dim k As Variant
k = "10"
MsgBox "2" < "-5"
End Sub
Public Function BracketInString(ByVal str As String, ByVal i As Integer) As Integer
    While i > 1
        i = i - 1
        If Mid(str, i, 1) = Chr(34) Then
            If i = 1 Then
                BracketInString = 0
                Exit Function
            End If
            If Mid(str, i - 1, 1) <> "\" Then
                BracketInString = 0
                Exit Function
            End If
        End If
    Wend
    BracketInString = 1
End Function
Public Function DoSplitObject(ByVal str As String, ByVal identifier As String) As String()
    Dim returnVal() As String
    Dim length As Integer
    Dim q, i, c, d, k, errorChecker As Integer
    c = 1
    d = 1
    k = 0
    length = Len(identifier)
    str = StringRefine.Trim(str)
    str = " " + str + identifier
    For i = 1 To Len(str)
        If Mid(str, i, length) = identifier Then
            If q = 0 Then
                c = i - 1
                k = k + 1
                ReDim Preserve returnVal(1000001)
                returnVal(k) = Mid(str, d, c - d + 1)
                d = i + length
            End If
        '' Kiem tra dau nhay kep trong chuoi co dinh kem dau \ hay khong
        ElseIf Mid(str, i, 1) = Chr(34) Then
             If q = 1 Then
                If Mid(str, i - 1, 1) <> "\" Then
                    q = 0
                    errorChecker = errorChecker + 1
                End If
            Else
                If Mid(str, i - 1, 1) <> "\" Then
                    errorChecker = errorChecker + 1
                End If
                q = 1
            End If
        End If
    Next i
    ReDim Preserve returnVal(1000001)
    returnVal(0) = CStr(k)
    If errorChecker Mod 2 <> 0 Then
        ReDim Preserve returnVal(1000001)
        returnVal(0) = "-1"
    End If
    DoSplitObject = returnVal
End Function
