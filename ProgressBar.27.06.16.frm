VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Email Saver Processing..."
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5745
   OleObjectBlob   =   "ProgressBar.27.06.16.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Activate()
    ' Set the width of the progress bar to 0.
    ProgressBar.labelProgress.Width = 0
    
    ProgressBar.AttachmentProgress.Width = 0
    ' Call the main subroutine.

    ' Call Main
    
    Call EmailSaver
End Sub




