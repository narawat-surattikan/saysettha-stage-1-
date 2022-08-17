Attribute VB_Name = "Saysettha"
Dim statement() As String
Dim line, maxFor, maxIf As Integer
Dim ForCount(1 To 1000000, 1 To 2), IfCount(1 To 1000000, 1 To 2) As Variant
Dim iLine As Integer

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Public Sub ErrorHandle()
If Slide1.errcount <> 0 Then Slide1.console = "Runtime Error. Check the code again."
End Sub
Public Sub ConsoleWrite(ByVal message As String)
Slide1.console = Slide1.console + message
End Sub
Public Sub CountFor()
    Dim v As Variant
    Dim ForCountReverse(1 To 1000000) As Integer
    Dim i, c1, c2 As Integer
    i = -1
    j = 0
    For Each v In statement
        i = i + 1
        If (Not v Like "//*") Then
            If v Like "for ?* :: ?* >> ?*" Then
                c1 = c1 + 1
                ForCount(c1, 1) = i
            End If
            If v Like "next*" Then
                c2 = c2 + 1
                ForCountReverse(c2) = i
            End If
        End If
    Next v
    If c1 <> c2 Then
        Slide1.errcount = Slide1.errcount + 1
    Else
        If c1 = 0 Then Exit Sub
        For i = c2 To 1 Step -1
            j = j + 1
            ForCount(j, 2) = ForCountReverse(i)
        Next
        For j = 1 To c1
            ''MsgBox ForCount(j, 1) & ">" & ForCount(j, 2)
        Next
        maxFor = c1
    End If
End Sub
Public Sub CountIf(ByVal strInput As String)
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~IfLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~IfLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~IfLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~IfLine:?*" Then
            s.Delete
        End If
    Next
    
    Dim v As Variant
    Dim stats() As String
    Dim stack_splited() As String
    Dim i, maxline, c1, c2 As Integer
    Dim stack, fixed, last_line As String
    stats = Split(strInput, vbNewLine)
    maxline = ArrayLen(stats) - 1
    
    For i = 0 To maxline
        fixed = StringRefine.Trim(stats(i))
        If StringRefine.CustomRegexChecker(fixed, "^if[ ]*[ ]*.+$") = 0 Then
            Dim shp As Object
            Set shp = ActivePresentation.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, -100, 50, 50)
            shp.Name = "$$Saysettha~~IfLine:" & CStr(i)
            shp.AlternativeText = ","
            stack = stack & CStr(i) & ","
        ElseIf StringRefine.CustomRegexChecker(fixed, "^elseif[ ]*[ ]*.+$") = 0 Or StringRefine.CustomRegexChecker(fixed, "^else[ ]*(::)?[ ]*$") = 0 Then
            If stack <> "" Then
                stack_splited = Split(stack, ",")
                last_line = stack_splited(ArrayLen(stack_splited) - 2)
                Slide1.Shapes("$$Saysettha~~IfLine:" & last_line).AlternativeText = _
                Slide1.Shapes("$$Saysettha~~IfLine:" & last_line).AlternativeText & CStr(i) & ","
            End If
        ElseIf StringRefine.CustomRegexChecker(fixed, "^endif[ ]*$") = 0 Then
            If stack <> "" Then
                stack_splited = Split(stack, ",")
                last_line = stack_splited(ArrayLen(stack_splited) - 2)
                Slide1.Shapes("$$Saysettha~~IfLine:" & last_line).TextFrame2.TextRange.Text = CStr(i)
                stack = Left(stack, Len(stack) - 1)
                While (Mid(stack, Len(stack), 1) <> ",")
                    stack = Left(stack, Len(stack) - 1)
                    If stack = "" Then
                        GoTo ExitWhileStack
                    End If
                Wend
ExitWhileStack:
            End If
        End If
    Next
    If stack <> "" Then
        Saysettha.ConsoleWrite ("if must go with end if in the following line: " & replace(stack, ",", " ") & "." & vbNewLine)
    End If
End Sub

Public Sub CountWhile(ByVal strInput As String)
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~WhileLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~WhileLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~WhileLine:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~WhileLine:?*" Then
            s.Delete
        End If
    Next
    
    Dim v As Variant
    Dim stats() As String
    Dim stack_splited() As String
    Dim i, maxline, c1, c2 As Integer
    Dim stack, fixed, last_line As String
    stats = Split(strInput, vbNewLine)
    maxline = ArrayLen(stats) - 1
    
    For i = 0 To maxline
        fixed = LCase(StringRefine.Trim(stats(i)))
        If StringRefine.CustomRegexChecker(fixed, "^while[ ]*.+$") = 0 Then
            Dim shp As Object
            Set shp = ActivePresentation.Slides(1).Shapes.AddShape(msoShapeRectangle, 0, -50, 50, 50)
            shp.Name = "$$Saysettha~~WhileLine:" & CStr(i)
            stack = stack & CStr(i) & ","
        ElseIf StringRefine.CustomRegexChecker(fixed, "^wend[ ]*$") = 0 Then
            If stack <> "" Then
                stack_splited = Split(stack, ",")
                last_line = stack_splited(ArrayLen(stack_splited) - 2)
                Slide1.Shapes("$$Saysettha~~WhileLine:" & last_line).TextFrame2.TextRange.Text = CStr(i)
                stack = Left(stack, Len(stack) - 1)
                While (Mid(stack, Len(stack), 1) <> ",")
                    stack = Left(stack, Len(stack) - 1)
                    If stack = "" Then
                        GoTo ExitWhileStack
                    End If
                Wend
ExitWhileStack:
            End If
        End If
    Next
    If stack <> "" Then
        Saysettha.ConsoleWrite ("if must go with end if in the following line: " & replace(stack, ",", " ") & "." & vbNewLine)
    End If
End Sub

Public Sub TestCountIf()
    Call CountIf(Slide1.editor.Text)
End Sub
Public Function GetNode(ByVal nodeArray As String, ByVal lineStatement As Integer, ByVal typeReturn As String)
    If nodeArray = "for" Then
        For i = 1 To maxFor
            If ForCount(i, 1) = lineStatement Then
                If typeReturn = "end" Then
                    GetNode = ForCount(i, 2)
                Else
                    GetNode = i
                End If
                Exit Function
            End If
        Next
    Else
    End If
End Function
Public Sub resetVariables()
    Dim s As Shape
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
    For Each s In Slide1.Shapes
        If s.Name Like "$$Saysettha~~Variables:?*" Then
            s.Delete
        End If
    Next
End Sub

Public Function WhileNavigation(ByVal whileStart As String, ByVal whileEnd As String)
    Dim Condition As String
    Dim ErrorHandler As String
    Dim i As Integer
    Condition = statement(whileStart)
    Condition = StringRefine.Trim(Condition)
    Condition = Right(Condition, Len(Condition) - 5)
    Condition = StringRefine.Trim(Condition)
    If Check(Condition) = 2 Then
        WhileNavigation = "e"
        Exit Function
    End If
    If whileStart + 1 = whileEnd Then
        WhileNavigation = whileEnd
        Exit Function
    End If
    whileStart = whileStart + 1
    While Check(Condition) = 0
        ErrorHandler = ProcessInitialized(whileStart, whileEnd - 1)
        If ErrorHandler <> "" Then
            WhileNavigation = ErrorHandler
            Exit Function
        End If
    Wend
    WhileNavigation = whileEnd
    Exit Function
End Function

Public Function ifNavigation(ByVal ifStart As String, ByVal ifList As String, ByVal ifEnd As String)
    Dim Condition As String
    Dim ErrorHandler, ifStack, numberNode As String
    Dim ifStackList() As String
    Dim i, k, ErrorCode As Integer
    Condition = statement(ifStart)
    Condition = StringRefine.Trim(Condition)
    Condition = Right(Condition, Len(Condition) - 2)
    Condition = StringRefine.Trim(Condition)
    
    ErrorCode = Check(Condition)
    ifStack = ifList
    
    If ErrorCode = 2 Then
        ifNavigation = "e"
        Exit Function
    End If
    
    If ifStart + 1 = ifEnd Then
        ifNavigation = ifEnd
        Exit Function
    End If
    
    ifStart = ifStart + 1
    If ErrorCode = 0 Then
        If ifStack = "," Then
            ErrorHandler = ProcessInitialized(ifStart, ifEnd - 1)
            If ErrorHandler <> "" Then
                ifNavigation = ErrorHandler
                Exit Function
            End If
            ifNavigation = ifEnd
            Exit Function
        Else
            ifStack = DeleteString(ifStack, 1, 1)
            For i = 1 To Len(ifStack)
                If Mid(ifStack, i, 1) = "," Then
                    k = i - 1
                    Exit For
                End If
            Next
            ErrorHandler = ProcessInitialized(CInt(ifStart), CInt(Mid(ifStack, 1, k)) - 1)
            If ErrorHandler <> "" Then
                ifNavigation = ErrorHandler
                Exit Function
            End If
            ifNavigation = ifEnd
            Exit Function
        End If
    Else
        
        While ifStack <> ","
            ifStackList = Split(ifStack, ",")
            If ArrayLen(ifStackList) = 3 Then
                Condition = statement(CInt(ifStackList(1)))
                Condition = StringRefine.Trim(Condition)
                If StringRefine.CustomRegexChecker(Condition, "^else[ ]*$") = 0 Then
                    GoTo ProcessDone1
                End If
                Condition = Right(Condition, Len(Condition) - 6)
                Condition = StringRefine.Trim(Condition)
                
                
                ErrorCode = Check(Condition)
                If ErrorCode = 2 Then
                    ifNavigation = "e"
                    Exit Function
                End If
                
                If ErrorCode = 0 Then
ProcessDone1:
                    If CInt(ifStackList(1)) + 1 = ifEnd Then
                        ifNavigation = ifEnd
                        Exit Function
                    End If
                    ErrorHandler = ProcessInitialized(CInt(ifStackList(1)) + 1, ifEnd - 1)
                    If ErrorHandler <> "" Then
                        ifNavigation = ErrorHandler
                        Exit Function
                    End If
                    ifNavigation = ifEnd
                    Exit Function
                End If
                
            Else
                Condition = statement(CInt(ifStackList(1)))
                Condition = StringRefine.Trim(Condition)
                If StringRefine.CustomRegexChecker(Condition, "^else[ ]*$") = 0 Then
                    GoTo ProcessDone2
                End If
                
                Condition = Right(Condition, Len(Condition) - 6)
                Condition = StringRefine.Trim(Condition)
                
                
                ErrorCode = Check(Condition)
                If ErrorCode = 2 Then
                    ifNavigation = "e"
                    Exit Function
                End If
                
                If ErrorCode = 0 Then
ProcessDone2:
                    If CInt(ifStackList(1)) + 1 = ifEnd Then
                        ifNavigation = ifEnd
                        Exit Function
                    End If
                    ErrorHandler = ProcessInitialized(CInt(ifStackList(1)) + 1, ifStackList(2) - 1)
                    If ErrorHandler <> "" Then
                        ifNavigation = ErrorHandler
                        Exit Function
                    End If
                    ifNavigation = ifEnd
                    Exit Function
                End If
                
            End If
            
            ''Delete Stack
            ifStack = DeleteString(ifStack, 1, 1)
            While Mid(ifStack, 1, 1) <> ","
                ifStack = DeleteString(ifStack, 1, 1)
            Wend
            
        Wend
    End If
    Exit Function
HandleComplete:
    ifNavigation = ifEnd
    Exit Function
End Function

Public Sub BuildProject(ByVal codeInput As String)
    ' Xu ly cac so lieu  va reset cac thong so
    ''Slide1.errcount = 0
    Slide1.console = "" ''"Saysettha Pre-processed Language (C) 2017 - 2022" & vbNewLine
    Dim v As Variant
    statement = Split(codeInput, vbNewLine)
    line = ArrayLen(statement) - 1
    iLine = -1
    For Each v In statement
        iLine = iLine + 1
        statement(iLine) = StringRefine.Trim(v)
    Next v
    Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text = ","
    
    Call CountWhile(codeInput)
    Call CountIf(codeInput)
    Call resetVariables
    
    Dim kHandle As String
    kHandle = ProcessInitialized(0, line)
    If kHandle <> "" Then Report.ErrorHandle (kHandle)
    
    ConsoleWrite (vbNewLine & "Program exits with the 0 code . . .")
End Sub

Public Function ProcessInitialized(ByVal startIn As Integer, ByVal endIn As Integer) As String
    Dim getReturnOneTime As Variant
    Dim i As Integer
    
    Dim endIfLine, elseIfList, ErrorCode As String
    
    For i = startIn To endIn
        If Not statement(i) Like "//*" And Not statement(i) = "" Then
            If statement(i) Like "$?*" Then
                getReturnOneTime = Variables.DeclareVariable(statement(i))
                If getReturnOneTime = 2 Then
                    ProcessInitialized = ("Warning :: Compile Error: declaring variable [" & _
                    StringRefine.Trim(statement(i)) & "] in line " & CStr(i + 1) & " is invalid. ")
                    Exit For
                ElseIf getReturnOneTime = 3 Then
                    ProcessInitialized = ("Warning :: Compile Error: Math error [" & _
                    StringRefine.Trim(statement(i)) & "] in line " & CStr(i + 1) & " is invalid. ")
                    Exit For
                ElseIf getReturnOneTime Like "value_assign:?*" Then
                    kHandle = StringRefine.Trim(Split(statement(i), "=", 2)(1))
                    ProcessInitialized = ("Warning :: Compile Error: assigning with the [" & _
                    kHandle & "] value in line " & CStr(i + 1) & " is invalid. ")
                    Exit For
                End If
            ElseIf StringRefine.CustomRegexChecker(statement(i), "^die[ ]*[(]*[ ]*[)]*[ ]*$") = 0 Then
                ProcessInitialized = "die"
                Exit Function
            ElseIf StringRefine.CustomRegexChecker(statement(i), "^goto[ ]*\d+[ ]*$") = 0 Then
                If Navigation.Go(statement(i)) - 1 <> i Then i = Navigation.Go(statement(i)) - 2
            ElseIf StringRefine.CustomRegexChecker(statement(i), "^while[ ]*.+$") = 0 Then
                endline = Slide1.Shapes("$$Saysettha~~WhileLine:" & CStr(i)).TextFrame2.TextRange.Text
                ErrorCode = WhileNavigation(i, endline)
                If StringRefine.VerifyNumber(ErrorCode) = 0 Then
                    i = ErrorCode
                ElseIf ErrorCode = "die" Then
                    ProcessInitialized = "die"
                    Exit Function
                Else
                    ProcessInitialized = ErrorCode
                    Exit For
                End If
            ElseIf StringRefine.CustomRegexChecker(statement(i), "^if[ ]*.+$") = 0 Then
                endline = Slide1.Shapes("$$Saysettha~~IfLine:" & CStr(i)).TextFrame2.TextRange.Text
                elseIfList = Slide1.Shapes("$$Saysettha~~IfLine:" & CStr(i)).AlternativeText
                ErrorCode = ifNavigation(i, elseIfList, endline)
                If StringRefine.VerifyNumber(ErrorCode) = 0 Then
                    i = ErrorCode
                ElseIf ErrorCode = "die" Then
                    ProcessInitialized = "die"
                    Exit Function
                Else
                    ProcessInitialized = ErrorCode
                    Exit For
                End If
            ElseIf statement(i) Like "cout*" Then
                If StringRefine.CustomRegexChecker(statement(i), "^cout[ ]*<<[ ]*.*") = 0 Then
                    getReturnOneTime = Cout.Process(statement(i))
                    If getReturnOneTime = "quote" Then
                        ProcessInitialized = ("Warning :: Compile Error: missing expected quote(s) in [" & _
                        StringRefine.Trim(statement(i)) & "], line " & CStr(i + 1) & " is invalid. ")
                        Exit For
                    ElseIf getReturnOneTime Like "quote:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        ProcessInitialized = ("Warning :: Compile Error: invalid [" & _
                        kHandle & "] string in line " & CStr(i + 1) & ". ")
                        Exit For
                    ElseIf getReturnOneTime Like "variable:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        ProcessInitialized = ("Warning :: Compile Error: invalid [" & _
                        kHandle & "] variable in line " & CStr(i + 1) & ". ")
                        Exit For
                    ElseIf getReturnOneTime Like "matherr:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        ProcessInitialized = ("Warning :: Compile Error: math error in the [" & _
                        kHandle & "] expression, line " & CStr(i + 1) & ". ")
                        Exit For
                    ElseIf getReturnOneTime Like "syntax:?*" Then
                        kHandle = StringRefine.Trim(Split(getReturnOneTime, ":", 2)(1))
                        ProcessInitialized = ("Warning :: Compile Error: syntax error in the [" & _
                        kHandle & "] statement, line " & CStr(i + 1) & ". ")
                        Exit For
                    End If
                Else
                    ProcessInitialized = ("Warning :: Compile Error: calling the [" & _
                    StringRefine.Trim(statement(i)) & "] statement in line " & CStr(i + 1) & " is invalid. ")
                    Exit For
                End If
            End If
        End If
    Next
End Function

