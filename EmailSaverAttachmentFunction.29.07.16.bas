Attribute VB_Name = "EmailSaverAttachmentFunction"
Public Function SaveAttachments(myDocs As String, sName As String, sCount As Long, SelectionCount As Variant, Item As MailItem)


Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Attachment

Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderPath As String

Dim wordApp As Word.Application
Dim wordDoc As Word.Document
Dim pdfName As String

Dim killDoc As String
Dim aProgress As Single


' Get the path to your My Documents folder
strFolderPath = myDocs & "\"

i = 1
lngCount = Item.Attachments.Count


With ProgressBar
        .attachmentLabel = "Processing Attachments..."
        .attachmentFrame.Caption = Format(0, "0%")
        .AttachmentProgress.Width = 0 * _
                (.attachmentFrame.Width - 15)
    End With

Do While i <= lngCount

  strFile = Item.Attachments(i).fileName
             
            
            'deals with word docs and coverts them to pdf
                If Right(strFile, 3) = "doc" Then
                   sCount = sCount + 1
                    ' Debug.Print "1" & strFile
                    ' Combine with the path to the Temp folder.
                    strFile = strFolderPath & strFile
                    Debug.Print "strFolderPath = " & strFolderPath
                
                    ' Save the attachment as a file.
                    Item.Attachments(i).SaveAsFile strFolderPath & sName & ".00" & ".doc" 'strFile
    
                    'WORDMEUP:
        
                    Set wordApp = CreateObject("Word.Application")
                    Set wordDoc = wordApp.Documents.Open(strFolderPath & sName & ".00.doc")
                
                
                    pdfName = strFolderPath & sName & ".0" & sCount & ".pdf"
                
                    Debug.Print "pdfName = " & pdfName
                
                    wordDoc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:= _
                    wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
                    Item:=wdExportDocumentContent, IncludeDocProps:=True
                
                    Debug.Print "Doc Converted to PDF"
                    'close word
                    wordDoc.Close saveChanges:=False
                    wordApp.Quit
                    Set wordDoc = Nothing
                    Set wordApp = Nothing
                
                    Debug.Print "Word should not be running in task manager"
    
                    killDoc = strFolderPath & sName & ".00" & ".doc"
                    Debug.Print "KillDoc = " & strFolderPath & sName & ".00" & ".doc"
    
                    Dim fso
                    
                    Set fso = CreateObject("Scripting.FileSystemObject")
    
                        If fso.FileExists(killDoc) Then
    
                             'If file exists, It will delete the file from source location
                             fso.DeleteFile killDoc, True
                             Debug.Print "Doc File Deleted Successfully"
    
                             Set fso = Nothing
    
                        End If
        
                 End If
' deals with jpg
           
 '       If SelectionCount = 1 Then
            If Right(strFile, 3) = "jpg" And strFile <> "image001.jpg" And strFile <> "image002.jpg" Then 'image001 is the Level 3 Instacom logo which appears as an attatchment
                sCount = sCount + 1
              '  strFile = strFolderPath & strFile
             Item.Attachments(i).SaveAsFile strFolderPath & sName & ".0" & sCount & ".jpg" ' strFile
            End If
  '      End If
        
'deals with png

        If Right(strFile, 3) = "PNG" Or Right(strFile, 3) = "png" Then
            sCount = sCount + 1
            Item.Attachments(i).SaveAsFile strFolderPath & sName & ".0" & sCount & ".png" ' strFile
        End If


        
      ' deals with PDF Attachments...
        
        If Right(strFile, 3) = "pdf" Or Right(strFile, 3) = "PDF" Then
        
            If sName = "T02" Or sName = "H35" Then
                sCount = sCount + 1
                Item.Attachments(i).SaveAsFile strFolderPath & sName & ".0" & sCount & ".pdf"

                ' T02 and H35 expections - no need to sCount + 1 in Item.Attachements(i) statement due to email not being saved
            ElseIf Right(strFile, 35) <> "Data Key + Special Requirements.pdf" And Right(strFile, 33) <> "Vodafone Special Requirements.pdf" And strFile <> "A5.pdf" And strFile <> "P01.pdf" Then
               sCount = sCount + 1
              '  strFile = strFolderPath & strFile
                Item.Attachments(i).SaveAsFile strFolderPath & sName & ".0" & sCount & ".pdf" ' strFile
            End If
        End If
           
 
    Debug.Print i & " " & strFile
    Debug.Print "sCount is now " & sCount
    
    aProgress = ((i) / Item.Attachments.Count)

    If sCount > 0 Then
        With ProgressBar
            .attachmentFrame.Caption = Format(aProgress, "0%")
            .AttachmentProgress.Width = aProgress * _
                    (.attachmentFrame.Width - 15)
        End With
    End If
    
i = i + 1
Loop
Exit Function


GoTo ExitSub


ERROR_MESSAGE:
MsgBox "Error - Attachments not saved; check file paths"

ExitSub:

Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

       
End Function




