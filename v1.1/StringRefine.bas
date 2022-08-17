Attribute VB_Name = "StringRefine"


Public Function Ins(ByVal source As String, ByVal str As String, ByVal i As Integer) As String
    Ins = Mid(source, 1, i - 1) & str & Mid(source, i, Len(source) - i + 1)
End Function

Public Function DeleteString(ByVal str As String, ByVal s As Integer, ByVal l As Integer) As String
    Dim ss As String
    If Len(str) >= s And 1 <= s And s <= Len(str) Then
        ss = Left(str, s - 1)
        DeleteString = ss + Right(str, Len(str) - s - l + 1)
    End If
End Function

Public Function Trim(ByVal str As String) As String
    If Len(str) <> 0 Then
        While Mid(str, 1, 1) = " " Or Mid(str, 1, 1) = vbTab
            str = Right(str, Len(str) - 1)
        Wend
        If Len(str) = 0 Then Exit Function
        While Mid(str, Len(str), 1) = " " Or Mid(str, Len(str), 1) = vbTab
            str = Left(str, Len(str) - 1)
        Wend
        Trim = str
    Else
        Trim = ""
    End If
End Function

Public Function VerifyUserdefinedName(ByVal strIn As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = "^\w+$"
        If .test(strIn) = True Then
            VerifyUserdefinedName = 0
        Else
            VerifyUserdefinedName = 1
        End If
    End With
End Function
Public Function VerifyNumber(ByVal strIn As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = "^-?\d*[.]?[0-9]+$"
        If .test(strIn) = True Then
            VerifyNumber = 0
        Else
            VerifyNumber = 1
        End If
    End With
End Function
Public Function CustomRegexChecker(ByVal strIn As String, ByVal strPattern As String) As String
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Pattern = strPattern
        If .test(strIn) = True Then
            CustomRegexChecker = 0
        Else
            CustomRegexChecker = 1
        End If
    End With
End Function
Public Function VerifyStringValue(ByVal strIn As String)
    strIn = replace(strIn, "\" & Chr(34), "<replace>")
    If InStr(1, strIn, Chr(34)) <> 0 Then
        VerifyStringValue = 1
    Else
        VerifyStringValue = 0
    End If
End Function
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

Public Function CheckBracket(ByVal str As String, ByVal z As Integer) As Integer
    Dim st As String
    st = Mid(str, 1, i)
    Dim k As Integer
    k = 0
    For i = 1 To z
        If BracketInString(str, i) = 1 Then
            If Mid(str, i, 1) = "(" Then
                If i = 1 Then
                    k = k + 1
                    GoTo thefuck
                End If
                k = k + 1
            ElseIf Mid(str, i, 1) = ")" Then
                k = k - 1
            End If
        End If
    Next
thefuck:
    If k = 0 Then
        CheckBracket = 1
        Exit Function
    Else  ''chia het la khong nam trong chuoi, khong chia het la  nam trong chuoi
        CheckBracket = 0
        Exit Function
    End If
End Function

Public Function ReplaceStr(ByVal strInput As String, ByVal find As String, ByVal replace As String)
    Dim i As Integer
    For i = Len(strInput) To 1 Step -1
        If LCase(Mid(strInput, i, Len(find))) = find Then
            If StringRefine.BracketInString(strInput, i) = 1 Then
                strInput = DeleteString(strInput, i, Len(find))
                strInput = Ins(strInput, replace, i)
            End If
        End If
    Next
    ReplaceStr = strInput
End Function
Sub fk()
Dim k As String
k = InputBox("")
k = ReplaceStr(k, "<=", "{")
k = ReplaceStr(k, ">=", "}")
End Sub

