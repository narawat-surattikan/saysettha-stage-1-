Attribute VB_Name = "Navigation"
Public Function Go(ByVal strInput As String) As Integer
    strInput = StringRefine.Trim(strInput)
    strInput = Right(strInput, Len(strInput) - 4)
    strInput = StringRefine.Trim(strInput)
    Go = CInt(strInput)
End Function
