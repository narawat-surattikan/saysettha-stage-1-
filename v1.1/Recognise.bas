Attribute VB_Name = "Recognise"
Public Function RString(ByVal strInput As String) As String()
    Dim returnVal() As String
    Dim str As String
    strInput = StringRefine.Trim(strInput)
    If strInput Like Chr(34) & "*" & Chr(34) Then
        strInput = (Right(strInput, Len(strInput) - 1))
        strInput = (Left(strInput, Len(strInput) - 1))
    End If
    str = replace(strInput, "\" & Chr(34), "<>")
    If InStr(1, str, Chr(34)) <> 0 Then
        ReDim Preserve returnVal(1 To 2)
        returnVal(1) = "-1"
        RString = returnVal
        Exit Function
    End If
    ReDim Preserve returnVal(1 To 2)
    returnVal(1) = ""
    returnVal(2) = replace(replace(strInput, "\" & Chr(34), Chr(34)), "\n", vbNewLine)
    RString = returnVal
    Exit Function
End Function
Public Function RVal(ByVal strInput As String) As String()
    Dim returnVal() As String
    Dim str As String
    strInput = StringRefine.Trim(strInput)
    If strInput Like "$#*" Or InStr(1, strInput, " ") <> 0 _
    Or StringRefine.VerifyUserdefinedName(Right(strInput, Len(strInput) - 1)) = 1 Then
        ReDim Preserve returnVal(1 To 2)
        returnVal(1) = "-1"
        RVal = returnVal
        Exit Function
    End If
    strInput = (Right(strInput, Len(strInput) - 1))
    If Variables.CheckExist(strInput) = 1 Then
        Call DeclareVariable_complete(strInput, "")
    End If
    strInput = Variables.getVariableValue(strInput)
    ReDim Preserve returnVal(1 To 2)
    returnVal(1) = ""
    returnVal(2) = replace(replace(strInput, "\" & Chr(34), Chr(34)), "\n", vbNewLine)
    RVal = returnVal
    Exit Function
End Function
Sub a()
Dim k As String
k = InputBox("")
MsgBox RVal(k)(1)
MsgBox RVal(k)(2)
End Sub
