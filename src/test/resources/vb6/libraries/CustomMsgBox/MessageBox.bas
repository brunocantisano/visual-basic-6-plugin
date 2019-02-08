Attribute VB_Name = "MessageBox"
Option Explicit

Public Function CustomMessage(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    Dim p_MessageBox As CustomMsgBox.cCustomMsgBox
    Set p_MessageBox = New CustomMsgBox.cCustomMsgBox
    
    Dim index As Integer, aux As Integer
    index = 0
    aux = 0
    
    If Title = "" Then
        Title = App.Title
    End If
    CustomMessage = p_MessageBox.MsgBox(Prompt, Buttons, Title)
    
    Set p_MessageBox = Nothing
End Function

