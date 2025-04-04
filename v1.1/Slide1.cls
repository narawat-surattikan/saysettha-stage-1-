VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Slide1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Runcode()
    Call Saysettha.BuildProject(editor.Text)
End Sub




Private Sub CommandButton1_Click()
Call Check(TextBox1.Text)
End Sub

Private Sub console_Change()

End Sub

Private Sub console_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
console.SelStart = Len(console.Text)
End Sub

Private Sub editor_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Shapes("Position").TextFrame2.TextRange.Text = " Cursor Position: " & editor.SelStart
End Sub

Private Sub editor_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Shapes("Position").TextFrame2.TextRange.Text = " Cursor Position: " & editor.SelStart
End Sub

Private Sub errcount_Change()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub
