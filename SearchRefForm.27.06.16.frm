VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchRefForm 
   Caption         =   "Enter Search Ref"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   OleObjectBlob   =   "SearchRefForm.27.06.16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchRefForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cancelButtonclicked As Boolean

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then CommandButton1_Click
 End Sub



Private Sub CancelButton1_Click()
    cancelButtonclicked = True
    Unload Me
End Sub

Private Sub CommandButton1_Click()
    If TextBox1.Value = "" And OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False And OptionButton4.Value = False And OptionButton5.Value = False Then

        MsgBox "Error -  Choose Terriroty and Enter Search Ref", vbCritical, "Input Error!"
    
    ElseIf OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False And OptionButton4.Value = False And OptionButton5.Value = False Then
         MsgBox "Error - No Territory selected", vbCritical, "Input Error!"

    ElseIf TextBox1.Value = "" Then

        MsgBox "Error - No Search Ref Entered", vbCritical, "Input Error!"

    ElseIf Len(TextBox1.Value) > 6 Then
                MsgBox "Search Ref is too long", vbCritical, "Input Error!"
    ElseIf Len(TextBox1.Value) < 6 Then
                MsgBox "Search Ref is too short", vbCritical, "Input Error!"
    Else
        Unload Me
        
    End If
End Sub



Private Sub OptionButton3_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    MsgBox "Please choose/input an option or click the 'Cancel' button!", vbCritical, "Alert!!"
  End If
End Sub
