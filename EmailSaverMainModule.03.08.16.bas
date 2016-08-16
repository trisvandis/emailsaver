Attribute VB_Name = "EmailSaverMainModule"

Sub StartEmailSaver()
    ProgressBar.Show
End Sub


'Email saver v3.3 by Tristan Bowles

Sub EmailSaver()
On Error GoTo ERROR_ALERT


Dim Selection As Selection
Dim obj As Object
Dim Item As MailItem

Set Selection = Application.ActiveExplorer.Selection

Dim myDocs As String
Dim UserRef As String
Dim FolderRef As String
Dim sName As String

Dim SelectionCount As Variant
Dim sCount As Long

Dim eArray As Variant
ReDim eArray(1 To 1) As String
Dim AcitveUserArray As Variant

Dim finalMessage As String

Dim killDoc As String
Dim strFile As String
Dim QuestionToMessageBox As String

Dim PctDone As Single

Dim Counter As Integer

Dim EmailSubject As String
Dim EmailSender As String

Dim MultipleEmailSaveTest As Boolean
Dim T11Affected As Boolean
Dim CheckFolderExists As Boolean

Dim StartTime As Double
Dim SecondsElapsed As Double

    StartTime = Timer
    
    For Each obj In Selection
   
    Set Item = obj
    
    Dim fso As Object, TmpFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    sCount = 0
    Counter = Counter + 1
    SelectionCount = Selection.Count
    MultipleEmailSaveTest = True
    T11Affected = False
    CheckFolderExists = False
    
    ProgressBar.labelProgress.Width = 0
    
    EmailSubject = Item.Subject 'send subject to email form
    EmailSender = Item.Sender
    
    'update progressbar
    With ProgressBar
        .Caption = "Saving Email " & Counter & " of " & SelectionCount
        .ProgressLabel1 = "Determining utility type, user ref and folder name..."
    End With
    
' ****SECTION 1 - Determine Folder name (FolderRef), user reference (UserRef) and the Utility type (sName)****

        Debug.Print "Section 1 - Print Email and Save Attachments"
   '1.1 - get folder reference from email subject title
        Select Case True
            Case EmailSubject Like "*LNW1*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "LNW1"), 9)
            Case EmailSubject Like "*LNE1*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "LNE1"), 9)
            Case EmailSubject Like "*SET1*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "SET1"), 9)
            Case EmailSubject Like "*SCT1*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "SCT1"), 9)
            Case EmailSubject Like "*WES1*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "WES1"), 9)
            Case EmailSubject Like "*LNW2*"    'future profing for when codes change to 20000's
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "LNW2"), 9)
            Case EmailSubject Like "*LNE2*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "LNE2"), 9)
            Case EmailSubject Like "*SET2*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "SET2"), 9)
            Case EmailSubject Like "*SCT2*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "SCT2"), 9)
            Case EmailSubject Like "*WES2*"
                FolderRef = Mid(EmailSubject, InStr(EmailSubject, "WES2"), 9)
            Case Else
                GoTo MANUAL_INPUT
        End Select
        
    '1.2 - get user reference

                getActiveUser UserRef, FolderRef, myDocs   'get active user array
                
                Debug.Print "User found: " + UserRef
                
                    If UserRef <> "No Match" Then
                        GoTo GET_UTILITY_CODE
                    End If
   
                    If SelectionCount = 1 Then
                        Debug.Print "Can't find user reference!"
                        manualInput UserRef, EmailSubject
                            If UserRef = "Cancelled" Then  'if cancel button is pressed, userref is cancelled = close sub
                                GoTo QUICK_CLOSE
                            End If
                        're-check if filepath exists
                        myDocs = "O:\Buried Services\BSRM\" & UserRef & "\" & FolderRef
                            If FolderExists(myDocs) Then
                                Debug.Print "Folderpath exists = " & myDocs
                                CheckFolderExists = True
                                GoTo GET_UTILITY_CODE
                            Else
                                Debug.Print "Can't find user reference!"
                                CheckFolderExists = False
'                                MultipleEmailSaveTest = False
'                                    With Item
'                                        .UnRead = False
'                                        .FlagIcon = olRedFlagIcon
'                                        .Categories = "Unable to save as multiple selection"
'                                        .Save
'                                    End With
                                 GoTo CONFIRMATION_MESSAGE
                            End If
                    End If
 
 '1.3 - get utility ref from email address
GET_UTILITY_CODE:
    Select Case True
        Case Item.To = "BS_Transmittals"
            sName = "B01"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress = "Plantenquiries@instalcom.co.uk"
            sName = "T02"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*atkinsglobal*"
            sName = "T03"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*catelecomuk*"
            sName = "T09"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress = "NRSWA@sky.uk"
            sName = "T11"
                If Item.Body Like "*will not*" Then
                    T11Affected = False
                Else
                    T11Affected = True
                End If
            Debug.Print "T11 Affected = " & T11Affected
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress = "osp-team@uk.verizon.com"
            sName = "T24"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*bbmmjv*"
            sName = "TMP17"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*TfL*"
            sName = "TMP10"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*bournemouthwater*"
            sName = "H05"
            GoTo FOLDER_CHECK
        Case Item.SenderEmailAddress Like "*southeastwater*"
            sName = "H35"
            GoTo FOLDER_CHECK
        Case Item.Body Like "*Cambridge Water Company Plc*"
            sName = "H09"
            GoTo FOLDER_CHECK
        Case Item.To <> "BS_Transmittals"
            GoTo MANUAL_INPUT 'goes to the manual entry option for individual selections
        End Select

FOLDER_CHECK:
'update progressbar
    PctDone = 20 / 100
    UpdateProgressBar PctDone
    ProgressBar.ProgressLabel1 = "Checking if file path exists..."
    'check to see folder exists, if it does, skip section 2
        If FolderExists(myDocs) Then
            Debug.Print "Folderpath exists = " & myDocs
            CheckFolderExists = True
            GoTo T02_EXCEPTION
        Else
            CheckFolderExists = False
        End If
 ' ****SECTION 2 - Manual Entry****
MANUAL_INPUT:
        Debug.Print "Section 1 - Complete!"
        Debug.Print "Section 2 - Manual Input"
        '2.1.1 - determine if multiple selections have been made. If so, skip manual entry forms and flag email

        If SelectionCount > 1 Then
            Debug.Print "Can't find folder, skipping email save due to multiple selections"
            MultipleEmailSaveTest = False
              
            If CheckFolderExists = False Then
                
                If EmailSubject Like "*OP*" Then  'OP Exception
                    Item.UnRead = True
                    GoTo CONFIRMATION_MESSAGE
                End If
                
                Item.Categories = "Can't save: Check folder exists!"
            
            Else
                Item.Categories = "Unable to save as multiple selection"
            End If
                
                With Item
                   .UnRead = False
                   .FlagIcon = olRedFlagIcon
                   .Save
                End With
            GoTo CONFIRMATION_MESSAGE
        
        ElseIf FolderExists(myDocs) Then
        '2.1.2 - if only a single selection is made, check if folderpath exists. If it does, call manaually utility form to get the utility code
            CheckFolderExists = True
            Debug.Print "Folderpath exists = " & myDocs & " needs utility code..."
          
            ManualUtility sName, EmailSubject, EmailSender
                If sName = "Cancelled" Then
                    GoTo QUICK_CLOSE
                End If
        Else
        '2.1.3 - If a single selection is made and folderpath doesnt exist,begin the manual entry form process
            Debug.Print "Can't find folder"
                
        'step 2.2.1 - manual entry process - find the folder
                
                GetSearchID FolderRef, EmailSubject
                        If FolderRef = "Cancelled" Then
                            GoTo QUICK_CLOSE
                        End If
                        PctDone = 30 / 100
                        UpdateProgressBar PctDone
        
        'step 2.2.2  - find the user
                   
                getActiveUser UserRef, FolderRef, myDocs   'get active user array
                        
                Debug.Print "User found: " + UserRef
                
                    If UserRef = "No Match" Then 'if not a match, call the manual entry form
                        manualInput UserRef, EmailSubject
                            If UserRef = "Cancelled" Then  'if cancel button is pressed, userref is cancelled = close sub
                                GoTo QUICK_CLOSE
                            End If
                    
                    PctDone = 40 / 100
                    UpdateProgressBar PctDone
                    End If

        'step 2.2.3 - call the utility ref manual entry form

                    If sName = "" Then
                        ManualUtility sName, EmailSubject, EmailSender
                            If sName = "Cancelled" Then
                                GoTo QUICK_CLOSE
                            End If
                    End If

                Debug.Print "Manually Input - User Ref = " + UserRef
                Debug.Print "Manually Input - Folder Ref = " + FolderRef
                Debug.Print "Manually Input- Utility Ref = " + sName

        're-check if filepath exists

            myDocs = "O:\Buried Services\BSRM\" & UserRef & "\" & FolderRef

                If FolderExists(myDocs) Then
                    CheckFolderExists = True
                    Debug.Print "Folderpath exists = " & myDocs
                    GoTo T02_EXCEPTION
                Else
        'error messages that break the sub
                    ProgressBar.ProgressLabel1 = "Error!!!!!"
                    MsgBox "Please check the user folder exists and try again.", vbCritical, "Can't Find Folder Path!"
                    GoTo QUICK_CLOSE
                    End If
        End If
         
   ' ****SECTION 3 - Print Email and save attachments****
            
T02_EXCEPTION:
        Debug.Print "Section 2 - Complete!"
        Debug.Print "Section 3 - Print Email and Save Attachments"
        Debug.Print "Folder Ref = " + FolderRef
        Debug.Print "User Ref = " + UserRef
        Debug.Print "Utility Ref = " + sName
        Debug.Print myDocs
        
        If sName = "T02" Then  'use this T02_EXCEPTION if you just need to save attachments and not the email
            GoTo ATTACHMENT_SAVER
        End If
        
        If sName = "H35" And sCount > 0 Then 'use this H35 if you just need to save attachments and not the email
            GoTo ATTACHMENT_SAVER
        End If
        
CREATE_WORD_APP:
    
    PctDone = 50 / 100
    UpdateProgressBar PctDone
    ProgressBar.ProgressLabel1 = "Printing Email..."
    
    PrintEmail Item, sName, myDocs, T11Affected 'makes a PDF of the email
    
ATTACHMENT_SAVER: 'determines is attachments need saving
    PctDone = 60 / 100
    UpdateProgressBar PctDone
    'update progressbar
    With ProgressBar
        .ProgressLabel1 = "Saving Attachments..."
    End With
    
    Select Case True
    Case sName = "B01"
        Debug.Print "No Attachments"
    Case sName = "T02" Or sName = "H35" Or sName = "T03"
        SaveAttachments myDocs, sName, sCount, SelectionCount, Item
    Case sName = "T09"
        SaveAttachments myDocs, sName, sCount, SelectionCount, Item
        If sCount >= 1 Then
                    killDoc = myDocs & "\" & sName & ".pdf"
                    strFile = myDocs & "\" & sName & ".03" & ".pdf"
                    Debug.Print "File Check = " & strFile
                    Debug.Print "KillDoc = " & killDoc
            
                If fso.FileExists(strFile) Then
                'If file exists, It will delete the file from source location
                    fso.DeleteFile killDoc, True
                    Debug.Print "Doc File Deleted Successfully"
                    killDoc = myDocs & "\" & sName & ".01" & ".pdf"
                    Name strFile As killDoc
                    MsgBox "Saved to folder " & UserRef & ", " & FolderRef & ", as T09.01 with " & sCount & " Attachments"
                    Else
                    MsgBox "Saved to folder " & UserRef & ", " & FolderRef & ", as T09 with " & sCount & " Attachments"
                 End If
            End If
    Case sName = "TMP10"
        SaveAttachments myDocs, sName, sCount, SelectionCount, Item
        killDoc = myDocs & "\" & sName & ".pdf"
        strFile = myDocs & "\" & sName & ".03.pdf"
        Debug.Print "File Check = " & fileCheck
        Debug.Print "KillDoc = " & killDoc
            If fso.FileExists(strFile) Then
            'If file exists, It will delete the file from source location
                fso.DeleteFile killDoc, True
                Debug.Print "Doc File Deleted Successfully"
                Name strFile As killDoc
    '        Set fso = Nothing
            End If
    Case sName <> "B01"
        SaveAttachments myDocs, sName, sCount, SelectionCount, Item
            If sCount > 0 And SelectionCount = 1 Then
           
                QuestionToMessageBox = MsgBox("Do you want to save the email?", vbYesNo, "Save email?")
            
                If QuestionToMessageBox = vbNo Then
                    killDoc = myDocs & "\" & sName & ".00.pdf"
                    fso.DeleteFile killDoc, True
                End If
                    

'                    If FileExists(strFile) Then
'                        fso.DeleteFile killDoc, True
'                        killDoc = myDocs & "\" & sName & ".01" & ".pdf"
'                        Name strFile As killDoc
'                        Debug.Print "Email deleted..."
'                    For i = 1 To 10
'                        killDoc = myDocs & "\" & sName & ".0" & i & ".pdf"
'                        strFile = myDocs & "\" & sName & ".0" & i + 1 & ".pdf"
'                        Debug.Print i & strFile & " loop dmn it"
'                        If FileExists(strFile) Then
'                            Name strFile As killDoc
'                        End If
'                    Next i
'                        GoTo END_SELECT
'                    End If
                
            End If

END_SELECT:
End Select
   
    If sCount = 0 Then 'rename the email if no attachments
        strFile = myDocs & "\" & sName & ".00.pdf"
        killDoc = myDocs & "\" & sName & ".pdf"
            If FileExists(strFile) Then
                Name strFile As killDoc
            End If
        ProgressBar.attachmentLabel = "No Attachments"
    End If
    
    
    If sCount > 0 Then 'rename the email if there is no .01.pdf
        
        strFile = myDocs & "\" & sName & ".01.pdf"
        killDoc = myDocs & "\" & sName & ".pdf"
        
        Select Case True
            Case FileExists(strFile)
                Debug.Print sName & ".01.pdf - File exists!"
            Case FileExists(killDoc)
                Debug.Print sName & ".pdf - File exists!"
            Case Else
        
            strFile = myDocs & "\" & sName & ".00.pdf"
            killDoc = myDocs & "\" & sName & ".02.pdf"

            If FileExists(strFile) Then
                    Debug.Print sName & ".00.pdf - File exists! **Rename file > " & sName & ".01**"""
                    killDoc = myDocs & "\" & sName & ".01.pdf"
                    Name strFile As killDoc
            ElseIf FileExists(killDoc) Then
                    Debug.Print sName & ".02.pdf - File exists!  **Rename file > " & sName & ".01**"
                    strFile = myDocs & "\" & sName & ".01.pdf"
                    Name killDoc As strFile
            End If
            End Select
    End If

ProgressBar.ProgressLabel1 = "Compiling confirmation report..."

PctDone = 100 / 100

UpdateProgressBar PctDone


  ' ****SECTION 4 - Confirmation Message and marking as complete****
        Debug.Print "Section 3 - Complete!"
        Debug.Print "Section 4 - Confirmation Message and marking as complete"
CONFIRMATION_MESSAGE:
ProgressBar.ProgressLabel1 = "Complete!"
        If SelectionCount = 1 Then 'display confirmation message if selection count is equal to 1
            
            SecondsElapsed = Round(Timer - StartTime, 2)
            
            
            If CheckFolderExists = True Then
                If sCount = 0 Then
                    MsgBox "Saved to folder " & UserRef & ", " & FolderRef & ", under code " & sName & ", with no attachments" & vbNewLine & "Completed in: " & SecondsElapsed & " seconds", vbInformation, "Put down your coffee, your email has saved!"
                Else
                    MsgBox "Saved to folder " & UserRef & ", " & FolderRef & ", under code " & sName & " with " & sCount & " attachment(s)" & vbNewLine & "Completed in: " & SecondsElapsed & " seconds", vbInformation, "Put down your coffee, your email has been saved!"
                End If
                
                    With Item
                        .UnRead = False
                        .FlagStatus = olFlagComplete
                        .Categories = ""
                        .Save
                    End With
            Else
                
                MsgBox "Unable to save: check folder exists" & vbNewLine & "Completed in: " & SecondsElapsed & " seconds", vbInformation, "Put down your coffee, your email has saved!"
                    With Item
                        .UnRead = False
                        .FlagIcon = olRedFlagIcon
                        .Categories = "Can't Save: Check folder exists!"
                        .Save
                    End With
            
            End If
            
        End If
        

        If SelectionCount > 1 Then  'display confirmation message if selection count is equal to 1
            
        Debug.Print MultipleEmailSaveTest
            If MultipleEmailSaveTest = False Then
                eArray(UBound(eArray)) = Chr(149) + " Failed to save: " & EmailSubject & " - please save individually"
                ReDim Preserve eArray(1 To UBound(eArray) + 1) As String
                
            Else

                eArray(UBound(eArray)) = Chr(149) + " " + UserRef & " > " & FolderRef & ", saved under code: " & sName & " with " & sCount & " attachments"
                ReDim Preserve eArray(1 To UBound(eArray) + 1) As String
                'insert array here

                With Item
                       
                    .UnRead = False
                    .FlagStatus = olFlagComplete
                    .Categories = ""
                    .Save
                End With

            End If
        End If
    
    
    PctDone = 100 / 100
    UpdateProgressBar PctDone
    With ProgressBar
        .ProgressLabel1 = "Complete!"
    End With
Debug.Print "Section 4 - Complete!"
CLOSING_PROCESS:
    Next obj
        
        

        If SelectionCount > 1 Then
            
            finalMessage = SelectionCount & " emails processed with the following results... " & vbNewLine
            
            For i = 1 To UBound(eArray)
                
                finalMessage = finalMessage & vbNewLine & eArray(i)
            Next i
        SecondsElapsed = Round(Timer - StartTime, 2)
        MsgBox finalMessage & "Completed in: " & SecondsElapsed & " seconds", vbInformation, "Confirmation Report"
     
     End If

QUICK_CLOSE:
Unload ProgressBar
Set obj = Nothing
Set Selection = Nothing
Set Item = Nothing

Exit Sub

ERROR_ALERT:
MsgBox "Error Alert!!! " & Err.Description
Resume CLOSING_PROCESS

End Sub


Sub UpdateProgressBar(PctDone As Single)
    With ProgressBar

        ' Update the Caption property of the Frame control.
        .frameProgress.Caption = Format(PctDone, "0%")

        ' Widen the Label control.
        .labelProgress.Width = PctDone * _
            (.frameProgress.Width - 15)

    End With

    ' The DoEvents allows the UserForm to update.
    DoEvents
End Sub



Function manualInput(UserRef As String, EmailSubject As String)


Dim frm As UserForm1


Set frm = New UserForm1


frm.SubjectLabel1 = EmailSubject


frm.Show vbModal
    If frm.cancelButtonclicked = True Then
        UserRef = "Cancelled"
    Else
        Select Case True
            Case frm.OptionButton1.Value
                UserRef = "TB5"
            Case frm.OptionButton2.Value
                UserRef = "CG5"
            Case frm.OptionButton3.Value
                UserRef = "FP5"
            Case frm.OptionButton4.Value
                UserRef = "JW8"
            Case frm.OptionButton5.Value
                UserRef = "CS6"
            Case frm.OptionButton6.Value
                UserRef = "MA5"
            Case frm.OptionButton8.Value
                UserRef = "KD7"
            Case frm.OptionButton9.Value
                 UserRef = "SF5"
            Case frm.OptionButton10.Value
                 UserRef = "TH5"
            Case frm.OptionButton11.Value
                UserRef = "EH5"
            Case frm.OptionButton12.Value
                UserRef = "MS5"
            Case frm.OptionButton13.Value
                 UserRef = "NR7"
            Case frm.OptionButton14.Value
                 UserRef = "NB5"
            Case frm.OptionButton15.Value
                UserRef = "SE8"
            Case frm.OptionButton16.Value
                UserRef = "JB5"
            Case frm.OptionButton17.Value
                UserRef = "TS6"
            Case frm.OptionButton18.Value
                UserRef = "LR6"
            Case frm.OptionButton19.Value
                UserRef = "RF6"
            Case frm.OptionButton20.Value
                UserRef = "TM6"
            Case frm.OptionButton21.Value
                UserRef = "PREPARED"
            Case frm.OptionButton22.Value
                UserRef = "JH5"
            Case frm.OptionButton23.Value
                UserRef = "RD5"
            Case frm.TextBox1.Value <> ""
                UserRef = frm.TextBox1.Value
        
        End Select
    End If
Unload frm

End Function


Function ManualUtility(sName As String, EmailSubject As String, EmailSender As String)

Dim frm As UtilityForm

Set frm = New UtilityForm

frm.SubjectLabel3 = "Subject: " & EmailSubject
frm.SenderLabel1 = "Sender: " & EmailSender

frm.Show vbModal

    If frm.cancelButtonclicked = True Then
        sName = "Cancelled"
    Else

        Select Case True
            Case frm.OptionButton1.Value
                sName = "P01"
            Case frm.OptionButton2.Value
                sName = "P02"
            Case frm.OptionButton3.Value
                sName = "P03"
            Case frm.OptionButton4.Value
                sName = "P04"
            Case frm.OptionButton5.Value
                sName = "P05"
            Case frm.OptionButton6.Value
                sName = "P06"
            Case frm.OptionButton8.Value
                sName = "P07A"
            Case frm.OptionButton9.Value
                 sName = "P08"
            Case frm.OptionButton10.Value
                 sName = "P09"
            Case frm.OptionButton11.Value
                sName = "P10A"
            Case frm.OptionButton12.Value
                sName = "P11"
            Case frm.OptionButton13.Value
                 sName = "P12"
            Case frm.OptionButton14.Value
                 sName = "P27"
            Case frm.OptionButton15.Value
                sName = "P13"
            Case frm.OptionButton16.Value
                sName = "P18"
            Case frm.OptionButton17.Value
                sName = "P15"
            Case frm.OptionButton18.Value
                sName = "P16"
            Case frm.OptionButton19.Value
                sName = "P18"
            Case frm.OptionButton20.Value
                sName = "P19"
            Case frm.OptionButton21.Value
                sName = "P20"
            Case frm.OptionButton22.Value
                sName = "P22"
            Case frm.OptionButton23.Value
               sName = "P23"
            Case frm.OptionButton24.Value
                sName = "P25"
            Case frm.OptionButton25.Value
                sName = "P26"
            Case frm.OptionButton26.Value
                sName = "P36"
            Case frm.OptionButton27.Value
                sName = "P37A"
            Case frm.OptionButton28.Value
                sName = "P28"
            Case frm.OptionButton29.Value
                 sName = "P29"
            Case frm.OptionButton30.Value
                sName = "P30A"
            Case frm.OptionButton31.Value
                sName = "P31"
            Case frm.OptionButton32.Value
                sName = "P32"
            Case frm.OptionButton33.Value
                 sName = "P38A"
            Case frm.OptionButton34.Value
                 sName = "P39"
            Case frm.OptionButton35.Value
                sName = "P40"
            Case frm.OptionButton36.Value
                sName = "P41"
            Case frm.OptionButton37.Value
                sName = "P42"
            Case frm.OptionButton38.Value
                sName = "P43"
            Case frm.OptionButton39.Value
                sName = "P44"
            Case frm.OptionButton40.Value
                sName = "P45"
            Case frm.OptionButton42.Value
                sName = "P51"
            Case frm.OptionButton43.Value
                sName = "T02"
            Case frm.OptionButton44.Value
                sName = "T03"
            Case frm.OptionButton45.Value
                sName = "T09"
            Case frm.OptionButton46.Value
                sName = "T11"
            Case frm.OptionButton47.Value
                sName = "T24"
            Case frm.OptionButton49.Value
                sName = "TMP10"
            Case frm.OptionButton50.Value
                sName = "TMP17"
            Case frm.OptionButton51.Value
                sName = "P55"
            Case frm.OptionButton52.Value
                sName = "P57"
            Case frm.OptionButton53.Value
                sName = "Z03"
            Case Else  ' no value chosen
        
        
                sName = frm.TextBox1.Value
        
                    If sName = "" Then
                            MsgBox "Not a valid file name", vbOKOnly, "Input Error!"
                    End If

        ' MsgBox "Please check folder exists and save manually.", vbOKOnly, "Can't Find Folder Path!"
        'GoTo ENDOFFUNCTION
        End Select
    End If
ENDOFFUNCTION:
Unload frm
End Function


Function GetSearchID(FolderRef As String, EmailSubject As String)
'This function gets the search ID from the user Input form

Dim folderRefPrefix As String
Dim frm As SearchRefForm

Set frm = New SearchRefForm
SHOW_FORM:

frm.SubjectLabel2 = EmailSubject
frm.Show vbModal


    If frm.cancelButtonclicked = True Then
        FolderRef = "Cancelled"
    Else
    
        Select Case True
            Case frm.OptionButton1.Value
                folderRefPrefix = "LNE"
            Case frm.OptionButton2.Value
                folderRefPrefix = "LNW"
            Case frm.OptionButton3.Value
                folderRefPrefix = "SET"
            Case frm.OptionButton4.Value
                folderRefPrefix = "WES"
            Case frm.OptionButton5.Value
                folderRefPrefix = "SCT"
        End Select
                
        FolderRef = folderRefPrefix & frm.TextBox1.Value
        
        Debug.Print "Folder ref is it working " & FolderRef

    End If
    
Unload frm

End Function


Function getActiveUser(UserRef As String, FolderRef As String, myDocs As String)

                Dim ActiveUserArray As Variant
                
               'add or remove user codes here...
                ActiveUserArray = Array("Prepared", "CG5", "CS6", "TB5", "JW8", "FP5", "EH5", "KD7", "SF5", "EM5", "RD5", "JH5", "TH5", "JB5", "SE8", "MS5", "LR6", "MA5", "NB5", "NR7", "RF6", "TM6", "TS6", "AX1", "RP7", "GM1", "MA1", "NS1", "JW7", "DF5")
                
                For i = 0 To UBound(ActiveUserArray)

                UserRef = ActiveUserArray(i)
                myDocs = "O:\Buried Services\BSRM\" & UserRef & "\" & FolderRef
                    If FolderExists(myDocs) And FolderRef <> "" Then
                        Debug.Print "Folderpath exists = " & myDocs
                        GoTo END_FUNCTION
                    Else
                      Debug.Print ActiveUserArray(i) & " is not a match " & myDocs
                    End If
                Next i
            UserRef = "No Match"
                
END_FUNCTION:
End Function

