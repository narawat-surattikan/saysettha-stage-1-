Attribute VB_Name = "Variables"
Public Function Confirm_Old(ByVal statement As String) As Integer
    If statement Like "$?*" Then
        If statement Like "$[! ]?*" Then
            statement = Right(statement, Len(statement) - 1)
            If InStr(1, statement, "=") = 0 Then
                If statement Like "?*[ ]*?*" Then
                    Confirm = 2
                    Exit Function
                Else
                    If InStr(1, statement, "(") = 0 And InStr(1, statement, ")") = 0 And Not statement Like "#*" Then
                        If StringRefine.VerifyUserdefinedName(statement) = 0 Then
                            Confirm = 0
                        Else
                            Confirm = 2
                        End If
                    Else
                        Confirm = 2
                    End If
                End If
            Else
                If Not statement Like "#*" Then
                    Confirm = 0
                Else
                    Confirm = 2
                End If
            End If
        Else
            Confirm = 2
        End If
    Else
        Confirm = 1
    End If
End Function
Public Function Confirm(ByVal statement As String) As Integer
    Confirm = 0
End Function
Public Function DeclareVariable(ByVal statement As String)
    ''MsgBox statement
    Dim removeDollar, refineName As String
    Dim twoDec() As String
    removeDollar = Right(statement, Len(statement) - 1)
    If InStr(1, statement, "=") = 0 Then
        If removeDollar Like "*[ ]*" Or removeDollar Like "#*" Or StringRefine.VerifyUserdefinedName(removeDollar) = 1 Then
            DeclareVariable = "2"
            Exit Function
        End If
        If CheckExist(removeDollar) = 1 Then
            Call DeclareVariable_complete(removeDollar, "")
        End If
    Else
        twoDec = Split(removeDollar, "=", 2)
        twoDec(0) = StringRefine.Trim(twoDec(0))
        twoDec(1) = StringRefine.Trim(twoDec(1))
        
        If Not twoDec(0) Like "#*" And Not StringRefine.VerifyUserdefinedName(twoDec(0)) = 1 Then
            '' Kiem tra ve trai
            If InStr(1, twoDec(0), " ") <> 0 Or twoDec(0) = "$" Then
                DeclareVariable = 2
                Exit Function
            End If
            If CheckExist(twoDec(0)) = 1 Then
                Call DeclareVariable_complete(twoDec(0), "")
            End If
            
             '' Check the rest of the assign process
            If twoDec(1) Like "$*" Then
                If InStr(1, twoDec(0), " ") <> 0 Or twoDec(1) = "$" Then
                    DeclareVariable = 2
                    Exit Function
                End If
                refineName = Right(twoDec(1), Len(twoDec(1)) - 1)
                If twoDec(1) Like "$*[ ]*" Or twoDec(1) Like "$#*" Or StringRefine.VerifyUserdefinedName(refineName) = 1 Then
                    DeclareVariable = "value_assign:" & twoDec(1)
                    Exit Function
                End If
                If CheckExist(refineName) = 1 Then
                    Call DeclareVariable_complete(refineName, "")
                End If
                Call resignVariable(twoDec(0), refineName, "$")
                Exit Function
            ElseIf twoDec(1) Like Chr(34) & "*" & Chr(34) Then
                refineName = twoDec(1)
                refineName = Right(refineName, Len(refineName) - 1)
                refineName = Left(refineName, Len(refineName) - 1)
                Call resignVariable(twoDec(0), refineName, "")
                Exit Function
            ElseIf StringRefine.VerifyNumber(twoDec(1)) = 0 Then
                refineName = twoDec(1)
                Call resignVariable(twoDec(0), refineName, "")
                Exit Function
            ElseIf StringRefine.CustomRegexChecker(LCase(twoDec(1)), "^calc[ ]*[(].+[)][ ]*$") = 0 Then
                refineName = twoDec(1)
                refineName = Right(refineName, Len(refineName) - 4)
                refineName = StringRefine.Trim(refineName)
                refineName = Right(refineName, Len(refineName) - 1)
                refineName = Left(refineName, Len(refineName) - 1)
                tempvalue = Calculator.Calc(refineName)
                If tempvalue <> "Math Error" Then
                    Call resignVariable(twoDec(0), tempvalue, "")
                Else
                    DeclareVariable = 3
                End If
                Exit Function
            ElseIf StringRefine.CustomRegexChecker(LCase(twoDec(1)), "^concat[ ]*[(].+[)][ ]*$") = 0 Then
                tempvalue = ConcatString(twoDec(1))(1)
                If tempvalue <> "e" Then
                    Call resignVariable(twoDec(0), ConcatString(twoDec(1))(2), "")
                Else
                    DeclareVariable = 3
                End If
                Exit Function
            Else
                DeclareVariable = "value_assign:" & twoDec(1)
                Exit Function
            End If
        Else
            DeclareVariable = 2
            Exit Function
        End If

        ''MsgBox Trim(twoDec(0)) & vbNewLine & Trim(twoDec(1))
    End If
End Function
Public Sub resignVariable(ByVal nameVariable1 As String, ByVal nameVariable2 As String, ByVal assignType As String)
    If assignType = "$" Then
        Slide1.Shapes("$$Saysettha~~Variables:" & nameVariable1).TextFrame2.TextRange.Text = _
        Slide1.Shapes("$$Saysettha~~Variables:" & nameVariable2).TextFrame2.TextRange.Text
    Else
        Slide1.Shapes("$$Saysettha~~Variables:" & nameVariable1).TextFrame2.TextRange.Text = nameVariable2
    End If
End Sub
Public Function replaceVar_Calc(ByVal str As String)
    Dim sp() As String
    str = Calculator.RefineInput(str, 1)
    sp = Split(str, " ")
    Dim v As Variant
    Dim val As String
    For Each v In sp
        If v Like "$*" Then
            val = Right(v, Len(v) - 1)
            If Not val Like "#*" And Not StringRefine.VerifyUserdefinedName(val) = 1 Then
                refineName = "0"
                If CheckExist(val) = 1 Then
                    Call DeclareVariable_complete(val, "")
                    Call resignVariable(val, "0", "")
                Else
                    If StringRefine.VerifyNumber(getVariableValue(val)) = 0 Then refineName = getVariableValue(val) Else _
                    refineName = "0"
                End If
                replaceVar_Calc = replaceVar_Calc + refineName + " "
            Else
                replaceVar_Calc = "error"
                Exit Function
            End If
        Else
            replaceVar_Calc = replaceVar_Calc + v + " "
        End If
    Next v
End Function

Public Function CheckExist(ByVal nameVariable As String) As Integer
    If InStr(1, Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text, _
    "," & "$$Saysettha~~Variables:" & nameVariable & ",") <> 0 Then
        CheckExist = 0
    Else
        CheckExist = 1
    End If
End Function
Public Function getVariableValue(ByVal nameVariable As String)
    getVariableValue = Slide1.Shapes("$$Saysettha~~Variables:" & nameVariable).TextFrame2.TextRange.Text
End Function

Public Sub DeclareVariable_complete(ByVal nameVariable As String, ByVal valueVariable As Variant)
    Dim shp As Object
    Set shp = ActivePresentation.Slides(1).Shapes.AddShape(msoShapeRectangle, -50, -50, 50, 50)
    shp.Name = "$$Saysettha~~Variables:" & nameVariable
    If valueVariable <> "" Then
        shp.TextFrame2.TextRange.Text = valueVariable
    End If
    Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text = _
    Slide1.Shapes("$$Saysettha~~VariablesStack").TextFrame2.TextRange.Text & "$$Saysettha~~Variables:" & nameVariable & ","
End Sub


