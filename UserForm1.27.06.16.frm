VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Enter User Ref"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8835
   OleObjectBlob   =   "UserForm1.27.06.16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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

cancelButtonclicked = False

Select Case True
    Case OptionButton1.Value
    Unload Me
    Case OptionButton2.Value
    Unload Me
    Case OptionButton3.Value
    Unload Me
    Case OptionButton4.Value
    Unload Me
    Case OptionButton5.Value
    Unload Me
    Case OptionButton6.Value
    Unload Me
    Case OptionButton8.Value
    Unload Me
    Case OptionButton9.Value
    Unload Me
    Case OptionButton10.Value
    Unload Me
    Case OptionButton11.Value
    Unload Me
    Case OptionButton12.Value
    Unload Me
    Case OptionButton13.Value
    Unload Me
    Case OptionButton14.Value
    Unload Me
    Case OptionButton15.Value
    Unload Me
    Case OptionButton16.Value
    Unload Me
    Case OptionButton17.Value
    Unload Me
    Case OptionButton18.Value
    Unload Me
    Case OptionButton19.Value
    Unload Me
    Case OptionButton20.Value
    Unload Me
    Case OptionButton21.Value
    Unload Me
    Case OptionButton22.Value
    Unload Me
    Case OptionButton23.Value
    Unload Me
    Case TextBox1.Value <> ""

        If Len(TextBox1.Value) > 8 Then
              MsgBox "User ref is too long", vbCritical, "Input Error!"
        ElseIf Len(TextBox1.Value) < 3 Then
                MsgBox "User ref is too short", vbCritical, "Input Error!"
        Else
            Unload Me
                
    End If
    
    Case Else
    MsgBox "No User selected - Please select or input a user reference", vbCritical, "Input Error!"
 End Select
 

End Sub




Private Sub OptionButton10_Click()

End Sub

Private Sub TextBox1_Change()

Dim oneControl As Object

For Each oneControl In Me.Controls

    If TypeName(oneControl) = "OptionButton" Then
        oneControl.Value = False
    End If

Next oneControl

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, _
  CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then
    Cancel = True
    MsgBox "Please choose/input an option or click the 'Cancel' button!", vbCritical, "Alert!!"
  End If
End Sub
