Attribute VB_Name = "Module1"


Sub DemoWhile()
Dim a, b As Integer
a = 1
b = 3
While a < 5
    Debug.Print "I hacked Project's face real\n"
    a = a + 1
Wend
End Sub

Private Function CheckConditionString(ByVal strInput As String, ByVal k As Integer) As String
    Dim i, c As Integer
    For i = k To Len(strInput)
        If Mid(strInput, i, 1) = "(" And StringRefine.BracketInString(strInput, i) = 1 Then
            c = c + 1
        ElseIf Mid(strInput, i, 1) = ")" And StringRefine.BracketInString(strInput, i) = 1 Then
            c = c - 1
        ElseIf (Mid(strInput, i, 1) = "=" Or Mid(strInput, i, 1) = "|") And StringRefine.BracketInString(strInput, i) = 1 Then
            CheckConditionString = 0
            Exit Function
        End If
    Next
    CheckConditionString = 1
    Exit Function
End Function
Private Sub subdffs()
    MsgBox CheckKeywordBefore("concat(125)", 7)
End Sub
Private Function CheckKeywordBefore(ByVal strInput As String, ByVal k As Integer)
    Dim i, c As Integer
    Dim che As String
    che = ""
    If k <= 0 Then
        CheckKeywordBefore = 1
        Exit Function
    End If
    For i = k To 1 Step -1
        If (StringRefine.VerifyUserdefinedName(Mid(strInput, i, 1)) = 0) And StringRefine.BracketInString(strInput, i) = 1 Then
            che = che & Mid(strInput, i, 1)
        Else
            Select Case Mid(strInput, i, 1)
                Case " ", "#", "{", "("
                    If che = "" Then
                        If CheckKeywordBefore(strInput, i - 1) = 1 Then
                            CheckKeywordBefore = 1
                            Exit Function
                        End If
                    ElseIf (StringRefine.VerifyUserdefinedName(che) = 0) And StringRefine.BracketInString(strInput, i) = 1 Then
                        CheckKeywordBefore = 0
                        Exit Function
                    End If
                Case "$"
                    If (StringRefine.VerifyUserdefinedName(che) = 0) And StringRefine.BracketInString(strInput, i) = 1 Then
                        CheckKeywordBefore = 1
                        Exit Function
                    End If
                Case "|", "&"
                    CheckKeywordBefore = 1
                    Exit Function
            End Select
        End If
    Next
End Function
Sub a1()
MsgBox CheckKeywordBefore("(((concat(fsdf,df)=1)))", 2)
End Sub
Private Function CheckConditionString2(ByVal strInput As String, ByVal i As Integer)
    Dim c As Integer
    If (Mid(strInput, i, 1) = "(" Or Mid(strInput, i, 1) = "{") And StringRefine.BracketInString(strInput, i) = 1 Then
        If CheckKeywordBefore(strInput, i - 1) = 1 Then
            CheckConditionString2 = 0
            Exit Function
        End If
    End If
    CheckConditionString2 = 1
    Exit Function
End Function
Function RefineConditionString(ByVal strInput As String) As String
    Dim i, last As Integer
    Dim stack As String
    stack = ","
    Dim k() As String
    strInput = TestRPNForCondition.RefineInput(strInput)
    For i = 1 To Len(strInput)
        If Mid(strInput, i, 1) = "(" And StringRefine.BracketInString(strInput, i) = 1 And CheckConditionString2(strInput, i) = 0 Then
            strInput = DeleteString(strInput, i, 1)
            strInput = Ins(strInput, "{", i)
        End If
    Next
    
    
    For i = 1 To Len(strInput)
        If (Mid(strInput, i, 1) = "{" Or Mid(strInput, i, 1) = "(") And StringRefine.BracketInString(strInput, i) = 1 Then
            stack = stack & CStr(i) & ","
        End If
        If (Mid(strInput, i, 1) = ")") And StringRefine.BracketInString(strInput, i) = 1 Then
            k = Split(stack, ",")
            last = CInt(k(ArrayLen(k) - 2))
            If Mid(strInput, last, 1) = "{" And StringRefine.BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
                strInput = Ins(strInput, "}", i)
            End If
            stack = Left(stack, Len(stack) - 1)
            While Mid(stack, Len(stack), 1) <> ","
                stack = Left(stack, Len(stack) - 1)
            Wend
        End If
    Next
    
    RefineConditionString = strInput
End Function
Sub a()
Debug.Print RefineConditionString("((((concat(1125,122)=1|((1=1)))))")

End Sub
