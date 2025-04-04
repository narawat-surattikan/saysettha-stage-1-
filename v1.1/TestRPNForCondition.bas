Attribute VB_Name = "TestRPNForCondition"

'Built-in Functions For Saysettha Language
'Copyright © 2017 – 2022 by Dongvan Technologies
'Do not change the code inside here if you dont know anything about it.
'We wont take any responsibilites for your modified code

Private Function Ins(ByVal source As String, ByVal str As String, ByVal i As Integer) As String
    Ins = Mid(source, 1, i - 1) & str & Mid(source, i, Len(source) - i + 1)
End Function

Private Function DeleteString(ByVal str As String, ByVal s As Integer, ByVal l As Integer) As String
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
    pt = Array("|", "&", "=", "`", "~", "<", ">", "!", "{", "}")
    For Each v In pt
        If v = str Then i = i + 1
    Next v
    If i = 0 Then InOperation = 1 Else InOperation = 0
End Function
Private Function STT(ByVal X As String) As Single
    'Check priority of an object.
    Select Case X
        Case "|", "&"
            STT = 1
        Case "=", ">", "<", "!", "`", "~"
            STT = 2
        Case "["
            STT = 0
    End Select
End Function
Public Sub refi()
MsgBox RefineInput("       (           (           (        concat  (     fsdf      ,     df       )    =   1   )   |    3   =    3    )    |   3   =   3  )    ")
End Sub
Public Function Calc(ByVal RefineInputString As String)
    ''Bracket handler
    Dim i, c, d As Integer
    On Error GoTo ErrorHandler
    RefineInputString = ConvertInput(RefineInput(RefineInputString))
    Calc = RefineInputString
    Exit Function
ErrorHandler:
    Calc = "Math Error"
End Function
Public Function ReplaceSpace(ByVal str As String) As String
    Dim i As Integer
    str = " " + str
    For i = 1 To Len(str)
        If Mid(str, i, 2) = "  " Then
            If BracketInString(str, i) = 1 Then
                str = StringRefine.DeleteString(str, i, 2)
                str = StringRefine.Ins(str, " ", i)
            End If
        End If
    Next i
    ReplaceSpace = str
End Function
Public Sub tess()
Dim k As String
k = InputBox("")
Debug.Print ReplaceSpace(k)
End Sub
Public Function BracketInString(ByVal str As String, ByVal i As Integer) As Integer
    Dim st As String
    st = Mid(str, 1, i)
    Dim k As Integer
    k = 0
    While i > 1
        i = i - 1
        If Mid(str, i, 1) = Chr(34) Then
            If i = 1 Then
                k = k + 1
                GoTo thefuck
            End If
            If Mid(str, i - 1, 1) <> "\" Then
                k = k + 1
            End If
        End If
    Wend
thefuck:
    If k Mod 2 = 0 Then
        If k <> 0 Then
            BracketInString = 1
            Exit Function
        Else
            BracketInString = 1
            Exit Function
        End If
    Else  ''chia het la khong nam trong chuoi, khong chia het la  nam trong chuoi
        BracketInString = 0
        Exit Function
    End If
End Function


Sub cal()
Slide1.TextBox2 = RefineInput(Slide1.TextBox1)
Slide1.TextBox2 = TestRPNForCondition.Calc(Slide1.TextBox1.Text)
Dim i As Variant
Dim k() As String
k = ConditionCheck.DoSplitObject(Slide1.TextBox2, " ")
Slide1.console.Text = ""
For i = 1 To CInt(k(0))
    Slide1.console.Text = Slide1.console.Text & k(i) & vbNewLine
Next
End Sub

Public Function ObjectInString(ByVal str As String, ByVal i As Integer, ByVal identifier As String) As Integer
    Dim st As String
    st = Mid(str, 1, i)
    Dim k As Integer
    k = 0
    While i > 1
        i = i - 1
        If Mid(str, i, 1) = identifier Then
            If i = 1 Then
                k = k + 1
                GoTo thefuck
            End If
            If BracketInString(str, i) = 1 Then
                ObjectInString = 0
                Exit Function
            End If
        End If
    Wend
thefuck:
    ObjectInString = 1
    Exit Function
End Function

Sub aaa()
Dim i As Variant
Dim k() As String
k = ConditionCheck.DoSplitObject(Slide1.TextBox2, " ")
For i = 1 To CInt(k(0))
    Debug.Print k(i)
Next
End Sub
Function RefineInput(ByVal strInput As String, Optional ref As Integer) As String
    'Handling
    Dim TempString As String
    Dim i, n As Integer
    strInput = strInput + " "
    For i = Len(strInput) - 1 To 1 Step -1
        If InOperation(Mid(strInput, i, 1)) = 0 Then
            If BracketInString(strInput, i) = 1 Then
                strInput = Ins(strInput, " ", i + 1)
                strInput = Ins(strInput, " ", i)
            End If
        Else
            If InOperation(Mid(strInput, i + 1, 1)) = 0 Then
                If BracketInString(strInput, i + 1) = 1 Then
                    strInput = Ins(strInput, " ", i)
                    strInput = Ins(strInput, " ", i + 1)
                End If
            End If
        End If
    Next i
    Do
        TempString = strInput
        strInput = ReplaceSpace(strInput)
    Loop Until TempString = strInput
    While (Mid(strInput, 1, 1) = " ")
        strInput = Right(strInput, Len(strInput) - 1)
    Wend
    
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = ", " Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i + 1, 1)
            End If
        End If
    Next i
    
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = " ," Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
            End If
        End If
    Next i
    
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = " )" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
            End If
        End If
    Next i
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = " (" Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, 1)
            End If
        End If
    Next i
    
    For i = Len(strInput) - 1 To 1 Step -1
        If (Mid(strInput, i, 2)) = "( " Then
            If BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i + 1, 1)
            End If
        End If
    Next i
    
    
        
    RefineInput = strInput
End Function

Public Function ConvertInput(ByRef strInput As String) As String
    On Error GoTo a:
    Dim t, stack, X As String
    Dim RPNString As String
    Dim i As Integer
    For i = 1 To Len(strInput) Step 1
         X = (Mid(strInput, i, 1))
         If X <> " " Then
              t = t + X
         Else
            If BracketInString(strInput, i) = 1 Then  'Chia het :: ngoai chuoi :: 1 khong chia het :: trong chuoi ::  0
                  Dim c, d As String
                  c = Mid(t, 1, 1)
                  If InOperation(c) = 0 Then
                       Select Case c
                       Case "{"
                            If BracketInString(strInput, i) = 0 Then
                                GoTo NextThing
                            Else
                                stack = stack + t
                            End If
                       Case "}"
                            If BracketInString(strInput, i) = 0 Then
                                GoTo NextThing
                            Else
                                Do
                                     d = Mid(stack, Len(stack), 1)
                                    stack = Left(stack, Len(stack) - 1)
                                     If d <> "{" Then
                                          RPNString = RPNString + d + " "
                                     Else
                                     Exit Do
                                     End If
                                Loop Until d = "}"
                            End If
                       Case Else
                            If BracketInString(strInput, i) = 0 Then
                                GoTo NextThing
                            Else
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
                            End If
                       End Select
                  Else
NextThing:
                       RPNString = RPNString + "" + t + " "
                  End If
                  t = ""
            Else
                RPNString = RPNString + "" + t + " "
                t = ""
            End If
         End If
    Next i
    While stack <> ""
         RPNString = RPNString + Mid(stack, Len(stack), 1) + " "
         stack = Left(stack, Len(stack) - 1)
    Wend
    Dim TempString As String
    strInput = RPNString
    Do
            TempString = strInput
            strInput = ReplaceSpace(strInput)
    Loop Until TempString = strInput
    If Mid(strInput, 1, 1) = " " Then ConvertInput = Right(strInput, Len(strInput) - 1) Else ConvertInput = strInput
    Exit Function
a:
ConvertInput = "Syntax Error"
End Function

Sub k()
Call Calculate(Slide1.TextBox2)
End Sub







