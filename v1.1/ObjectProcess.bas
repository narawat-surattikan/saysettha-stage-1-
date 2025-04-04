Attribute VB_Name = "ObjectProcess"
Public Function ConcatString(ByVal strInput As String) As String()
    Dim st() As String
    Dim jt() As String
    Dim oneTimeValue, twoTimeValue, conCatFullString, patternObject As String
    Dim i As Integer
    
    strInput = StringRefine.Trim(strInput)
    If StringRefine.CustomRegexChecker(LCase(strInput), "^concat[ ]*[(].+[)][ ]*$") = 1 Then
        ReDim Preserve st(1 To 2)
        st(1) = "e"
        ConcatString = st
        Exit Function
    End If
    strInput = StringRefine.Trim(Right(strInput, Len(strInput) - 6))
    strInput = (Right(strInput, Len(strInput) - 1))
    strInput = (Left(strInput, Len(strInput) - 1))
    strInput = StringRefine.Trim(strInput)
    jt = ObjectProcess.DoSplitObjectNew(strInput, ",")
    
    If jt(0) = "-1" Then
        ReDim Preserve st(1 To 2)
        st(1) = "e"
        ConcatString = st
        Exit Function
    End If
    
    For i = 1 To CStr(jt(0))
        patternObject = Chr(34) & "*" & Chr(34)
        If StringRefine.Trim(jt(i)) Like patternObject Then
            oneTimeValue = Recognise.RString(jt(i))(1)
            twoTimeValue = Recognise.RString(jt(i))(2)
            If oneTimeValue = "-1" Then
                ReDim Preserve st(1 To 2)
                st(1) = "e"
                st(2) = "quote"
                ConcatString = st
                Exit Function
            End If
            conCatFullString = conCatFullString & twoTimeValue
        ElseIf StringRefine.CustomRegexChecker(StringRefine.Trim(LCase(jt(i))), "^concat[ ]*[(].+[)][ ]*$") = 0 Then
            oneTimeValue = ConcatString(jt(i))(1)
            twoTimeValue = ConcatString(jt(i))(2)
            If oneTimeValue = "e" Then
                ReDim Preserve st(1 To 2)
                st(1) = "e"
                st(2) = "variable"
                ConcatString = st
                Exit Function
            End If
            conCatFullString = conCatFullString & twoTimeValue
        ElseIf StringRefine.Trim(jt(i)) Like "$*" Then
            oneTimeValue = Recognise.RVal(jt(i))(1)
            twoTimeValue = Recognise.RVal(jt(i))(2)
            If oneTimeValue = "-1" Then
                ReDim Preserve st(1 To 2)
                st(1) = "e"
                st(2) = "variable"
                ConcatString = st
                Exit Function
            End If
            conCatFullString = conCatFullString & twoTimeValue
        End If
    Next i

    ReDim Preserve st(1 To 2)
    st(2) = conCatFullString
    ConcatString = st
End Function

Private Sub test()
    k = InputBox("Input")
    MsgBox ConcatString(k)(2)
End Sub
Public Function DoSplitObjectNew(ByVal str As String, ByVal identifier As String) As String()
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
        If Mid(str, i, length) = identifier And StringRefine.CheckBracket(str, i) = 1 And StringRefine.BracketInString(str, i) = 1 Then
            c = i - 1
            k = k + 1
            ReDim Preserve returnVal(1000001)
            returnVal(k) = Mid(str, d, c - d + 1)
            d = i + length
        End If
    Next i
    ReDim Preserve returnVal(1000001)
    returnVal(0) = CStr(k)
    If errorChecker Mod 2 <> 0 Then
        ReDim Preserve returnVal(1000001)
        returnVal(0) = "-1"
    End If
    DoSplitObjectNew = returnVal
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


