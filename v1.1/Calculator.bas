Attribute VB_Name = "Calculator"

'Built-in Functions For Saysettha Language
'Copyright © 2017 – 2022 by Dongvan Technologies
'Do not change the code inside here if you dont know anything about it.
'We wont take any responsibilites for your modified code

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

Private Function InOperation(ByVal str As String, Optional a As Integer) As Integer
    Dim pt As Variant
    Dim v As Variant
    Dim i As Integer
    pt = Array("+", "-", "*", "/", "(", ")", "^", "|", "%", "—", "m")
    For Each v In pt
        If v = str Then i = i + 1
    Next v
    If i = 0 Then InOperation = 1 Else InOperation = 0
End Function
Private Function STT(ByVal X As String) As Single
    'Check priority of an object.
    Select Case X
        Case "—"
            STT = 5
        Case "gffgdfgfgdfgfgdfsg"
            STT = 3
        Case "^", "|"
            STT = 3
        Case "*", "/", "%", "d"
            STT = 2
        Case "+", "-"
            STT = 1
        Case "("
            STT = 0
    End Select
End Function

Public Function Calc(ByVal RefineInputString As String)
    ''Bracket handler
    Dim i, c, d As Integer
    On Error GoTo ErrorHandler
    RefineInputString = replace(RefineInputString, ",", ".")
    
    Dim StringRecheck As String
    StringRecheck = Variables.replaceVar_Calc(RefineInputString)
    If StringRecheck = "error" Then GoTo ErrorHandler
    RefineInputString = ConvertInput(StringRecheck)
    Calc = Calculate(RefineInputString)
    Exit Function
ErrorHandler:
    Calc = "Math Error"
End Function

Private Function HandleSubstring(ByVal strInput As String) As String
    Dim TempString As String
    Dim a, b, c As Boolean
    Dim i, n, k As Integer
    For i = Len(strInput) - 1 To 1 Step -1
        a = InOperation(Mid(strInput, i, 1)) = 0 And i = 1 And (Mid(strInput, i, 1) = "+" Or Mid(strInput, i, 1) = "-")
        b = InOperation(Mid(strInput, i, 1)) = 0 And (Mid(strInput, i, 1) = "+" Or Mid(strInput, i, 1) = "-") And (i - 2 > 0)
        If a Then
            n = i
            strInput = Ins(strInput, "0 ", n)
        End If
        If b Then
            If Mid(strInput, i - 2, 1) = "(" Then
                n = i
                strInput = Ins(strInput, "0 ", n)
            End If
        End If
    Next i
    Do
            TempString = strInput
            ' + -
            strInput = replace(strInput, " " & "+" & " " & "+" & " ", " " & "+" & " ")
            strInput = replace(strInput, " " & "-" & " " & "-" & " ", " " & "+" & " ")
            strInput = replace(strInput, " " & "+" & " " & "-" & " ", " " & "-" & " ")
            strInput = replace(strInput, " " & "-" & " " & "+" & " ", " " & "-" & " ")
            ' * + -
            strInput = replace(strInput, " " & "*" & " " & "+" & " ", " " & "* ( 1 - 0 ) *" & " ")
            strInput = replace(strInput, " " & "*" & " " & "-" & " ", " " & "* ( 1 - 2 ) *" & " ")
            ' / + -
            strInput = replace(strInput, " " & "/" & " " & "+" & " ", " " & "/ ( 1 - 0 ) *" & " ")
            strInput = replace(strInput, " " & "/" & " " & "-" & " ", " " & "/ ( 1 - 2 ) *" & " ")
            ' ) + -
            strInput = replace(strInput, " " & "+" & " " & ")" & " ", " " & ")" & " ")
            strInput = replace(strInput, " " & "-" & " " & ")" & " ", " " & ")" & " ")
    Loop Until TempString = strInput
    Do
            TempString = strInput
            strInput = replace(strInput, " " & "^" & " " & "+" & " ", " " & "^" & " ")
            strInput = replace(strInput, " " & "^" & " " & "-" & " ", " " & "^ —" & " ")

            strInput = replace(strInput, " " & "|" & " " & "+" & " ", " " & "s" & " ")
            strInput = replace(strInput, " " & "|" & " " & "-" & " ", " " & "s —" & " ")
    Loop Until TempString = strInput
    HandleSubstring = strInput
End Function
 Function RefineInput(ByVal strInput As String, Optional ref As Integer) As String

    'Replace Stuff
    strInput = replace(strInput, "sqrt", "|")
    'Handling
    Dim TempString As String
    Dim i, n As Integer
    strInput = strInput + " "
    For i = Len(strInput) - 1 To 1 Step -1
        If InOperation(Mid(strInput, i, 1)) = 0 Or InOperation(Mid(strInput, i + 1, 1)) = 0 Then
            strInput = Ins(strInput, " ", i + 1)
        End If
    Next i
    Do
        TempString = strInput
        strInput = replace(strInput, Space(2), Space(1))
    Loop Until TempString = strInput
    While (Mid(strInput, 1, 1) = " ")
        strInput = Right(strInput, Len(strInput) - 1)
    Wend
    If ref = 1 Then
        strInput = HandleSubstring(strInput)
    End If
    RefineInput = strInput
End Function

Private Function ConvertInput(ByRef strInput As String) As String
    On Error GoTo a:
    Dim t, stack, X As String
    Dim RPNString As String
    Dim i As Integer
    For i = 1 To Len(strInput) Step 1
         X = (Mid(strInput, i, 1))
         If X <> " " Then
              t = t + X
         Else
              Dim c, d As String
              c = Mid(t, 1, 1)
              If InOperation(c) = 0 Then
                   Select Case c
                   Case "("
                        stack = stack + t
                   Case ")"
                        Do
                             d = Mid(stack, Len(stack), 1)
                            stack = Left(stack, Len(stack) - 1)
                             If d <> "(" Then
                                  RPNString = RPNString + d + " "
                             Else
                             Exit Do
                             End If
                        Loop Until d = "("
                   Case Else
                        If Not stack = "" Then
                             While stack <> "" And STT(c) <= STT(Mid(stack, Len(stack), 1))
                                  RPNString = RPNString + Mid(stack, Len(stack), 1) + " "
                                  stack = Left(stack, Len(stack) - 1)
                                  If Len(stack) = 0 Then
                                     GoTo ExitLoops
                                  End If
                             Wend
                        End If
ExitLoops:
                        stack = stack + c
                   End Select
              Else
                   RPNString = RPNString + t + " "
              End If
              t = ""
         End If
    Next i
    While stack <> ""
         RPNString = RPNString + Mid(stack, Len(stack), 1) + " "
         stack = Left(stack, Len(stack) - 1)
    Wend
    RPNString = replace(RPNString, "qrt", "")
    Dim TempString As String
    strInput = RPNString
    Do
            TempString = strInput
            strInput = replace(strInput, Space(2), Space(1))
    Loop Until TempString = strInput
    If Mid(strInput, 1, 1) = " " Then ConvertInput = Right(strInput, Len(strInput) - 1) Else ConvertInput = strInput
    Exit Function
    
    ' Replace (+/-)
    'Dim reg As New RegExp
    
    
a:
ConvertInput = "Syntax Error"
End Function

' In the function style
Private Function Calculate(ByVal strInput As String)
On Error GoTo Handle
Dim str, t As String
Dim i, count, e As Integer
Dim st(0 To 10000000)
If Right(strInput, 1) <> " " Then
strInput = strInput + " "
End If
count = -1
For i = 1 To Len(strInput)
    If Mid(strInput, i, 1) <> " " Then
        t = t + Mid(strInput, i, 1)
    Else
        If InOperation(Mid(t, 1, 1)) = 1 Then
            count = count + 1
            st(count) = val(t)
        Else
            If Mid(t, 1, 1) = "+" Then
                st(count - 1) = st(count - 1) - -(st(count))
                st(count) = ""
            ElseIf Mid(t, 1, 1) = "-" Then
                st(count - 1) = st(count - 1) - (st(count))
                st(count) = ""
            ElseIf Mid(t, 1, 1) = "*" Then
                st(count - 1) = st(count - 1) * (st(count))
                st(count) = ""
            ElseIf Mid(t, 1, 1) = "/" Then
                If st(count) = 0 Then
                    Calculate = "Math Error"
                    Exit Function
                End If
                st(count - 1) = st(count - 1) / (st(count))
                st(count) = ""
            ElseIf Mid(t, 1, 1) = "^" Then
                st(count - 1) = st(count - 1) ^ (st(count))
                st(count) = ""
            ElseIf Mid(t, 1, 1) = "|" Then
                If st(count) < 0 Then
                    e = 1
                    Exit For
                Else
                    st(count) = Math.Sqr(st(count))
                    count = count + 1
                End If
            ElseIf Mid(t, 1, 1) = "—" Then
                st(count) = st(count) * -1
                count = count + 1
            ElseIf Mid(t, 1, 1) = "%" Then
                st(count - 1) = st(count - 1) Mod (st(count))
                st(count) = ""
            End If
            count = count - 1
        End If
        t = ""
    End If
Next i

If e = 0 And st(1) = "" Then
Calculate = st(0)
Else
Calculate = "Math Error"
End If
Exit Function
Handle:
Calculate = "Math Error"
End Function




