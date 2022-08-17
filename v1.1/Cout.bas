Attribute VB_Name = "Cout"
Dim k As String
Private Function Ins(ByVal source As String, ByVal str As String, ByVal i As Integer) As String
    Ins = Mid(source, 1, i - 1) & str & Mid(source, i, Len(source) - i + 1)
End Function

Private Function DeleteString(str As String, s As Integer, l As Integer) As String
    Dim ss As String
    If Len(str) >= s And 1 <= s And s <= Len(str) Then
        ss = Left(str, s - 1)
        DeleteString = ss + Right(str, Len(str) - s - l + 1)
    End If
End Function


Sub a()
'k = InputBox("Test", "Test") '' "cout < 'Gumball               Pokemon'"
'MsgBox "Result = '" & Process(k)(1) & "'"
Slide1.Shapes("console").OLEFormat.Object.Text = ""
End Sub

Public Function Process(ByVal str As String)
    Dim maxC As Integer
    Dim SplitObject() As String
    Dim SubObject, TotalString As String
    SplitObject = DoSplitObject(str)
    maxC = CInt(SplitObject(0))
    If maxC = "-1" Then
        Process = "quote"
        Exit Function
    End If
    For i = 1 To maxC
        SubObject = StringRefine.Trim(SplitObject(i))
        reg = "^[" & Chr(34) & "].*[" & Chr(34) & "]$"
        If StringRefine.CustomRegexChecker(SubObject, reg) = 0 Then
            oneTimeError = SubObject
            SubObject = Left(SubObject, Len(SubObject) - 1)
            SubObject = Right(SubObject, Len(SubObject) - 1)
            If StringRefine.VerifyStringValue(SubObject) = 1 Then
                Process = "quote:" & oneTimeError
                Exit Function
            Else
                TotalString = TotalString + InsertObjectString(SubObject) 'Replace(Replace(SubObject, "\" & Chr(34), Chr(34)), "\n", vbNewLine)
            End If
        ElseIf StringRefine.VerifyNumber(SubObject) = 0 Then
            TotalString = TotalString + InsertObjectString(SubObject)
        ElseIf SubObject Like "$*" Then
            If InStr(1, SubObject, " ") <> 0 Or _
            StringRefine.VerifyUserdefinedName(Right(SubObject, Len(SubObject) - 1)) = 1 Or _
            Right(SubObject, Len(SubObject) - 1) Like "#*" Then
                Process = "variable:" & SubObject
                Exit Function
            End If
            If Variables.CheckExist(Right(SubObject, Len(SubObject) - 1)) = 1 Then
                Call Variables.DeclareVariable_complete(Right(SubObject, Len(SubObject) - 1), "")
            End If
            TotalString = TotalString + InsertObjectString(Variables.getVariableValue(Right(SubObject, Len(SubObject) - 1)))
        ElseIf SubObject Like "calc*" Then
            If StringRefine.CustomRegexChecker(LCase(SubObject), "^calc[ ]*[(].+[)][ ]*$") = 0 Then
                refineName = SubObject
                refineName = Right(refineName, Len(refineName) - 4)
                refineName = StringRefine.Trim(refineName)
                refineName = Right(refineName, Len(refineName) - 1)
                refineName = Left(refineName, Len(refineName) - 1)
                tempvalue = Calculator.Calc(refineName)
                If tempvalue <> "Math Error" Then
                    TotalString = TotalString & tempvalue
                Else
                    Process = "matherr:" & SubObject
                    Exit Function
                End If
            Else
                Process = "syntax:" & SubObject
                Exit Function
            End If
        ElseIf SubObject Like "concat*" Then
            If StringRefine.CustomRegexChecker(LCase(SubObject), "^concat[ ]*[(].+[)][ ]*$") = 0 Then
                refineName = ObjectProcess.ConcatString(SubObject)(1)
                tempvalue = ObjectProcess.ConcatString(SubObject)(2)
                If refineName = "e" Then
                    Process = "syntax:" & SubObject
                    Exit Function
                End If
                TotalString = TotalString & tempvalue
            Else
                Process = "syntax:" & SubObject
                Exit Function
            End If
        End If
    Next
    Saysettha.ConsoleWrite (TotalString)
    Process = 0
End Function
Public Function InsertObjectString(ByVal SubObject As String)
    InsertObjectString = replace(replace(SubObject, "\" & Chr(34), Chr(34)), "\n", vbNewLine)
End Function

Public Function DoSplitObject(ByVal str As String) As String()
    Dim returnVal() As String
    ''ReDim Preserve returnVal(1000001)
    Dim q, i, c, d, k, errorChecker As Integer
    c = 1
    d = 3
    k = 0
    
    str = StringRefine.Trim(str)
    str = StringRefine.Trim(Right(str, Len(str) - 4)) + "<<"
     
    For i = 3 To Len(str)
        If Mid(str, i, 2) = "<<" Then
            If q = 0 Then
                c = i - 1
                k = k + 1
                ReDim Preserve returnVal(1000001)
                returnVal(k) = Mid(str, d, c - d + 1)
                d = i + 2
            End If
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
NextIteration:
    Next i
    ReDim Preserve returnVal(1000001)
    returnVal(0) = CStr(k)
    If errorChecker Mod 2 <> 0 Then
        ReDim Preserve returnVal(1000001)
        returnVal(0) = "-1"
    End If
    DoSplitObject = returnVal
End Function

'' Old function

Private Function Process_Old(ByVal str As String) As String()
    Dim returnVal(0 To 1000000) As String
    Dim q, i, c, d, k As Integer
    c = 1
    d = 2
    str = StringRefine.Trim(str)
    str = StringRefine.Trim(Right(str, Len(str) - 4)) + "<"
     
    For i = 2 To Len(str)
        If Mid(str, i, 1) = "<" Then
            If q = 0 Then
                c = i - 1
                k = k + 1
                returnVal(k) = Mid(str, d, c - d + 1)
                MsgBox returnVal(k)
                d = i + 1
            End If
        ElseIf Mid(str, i, 1) = Chr(34) Then
             If q = 1 Then
                If Mid(str, i - 1, 1) <> "\" Then
                    q = 0
                End If
            Else
                q = 1
            End If
        End If
NextIteration:
    Next i
    returnVal(0) = CStr(k)
    Process_Old = returnVal
End Function
