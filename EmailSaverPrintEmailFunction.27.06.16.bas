Attribute VB_Name = "EmailSaverPrintEmailFunction"
Function PrintEmail(Item As Object, sName As String, myDocs As String, T11Affected As Boolean)

Dim strToSaveAs As String
Dim strToSaveAs2 As String
Dim wrdApp As Word.Application
Dim wrdDoc As Word.Document


Set fso = CreateObject("Scripting.FileSystemObject")
Set tmpFileName = fso.GetSpecialFolder(2)

Set wrdApp = CreateObject("Word.Application")

ReplaceCharsForFileName sName, "-"
tmpFileName = tmpFileName & "\" & sName & ".mht"

Item.SaveAs tmpFileName, olMHTML

Set wrdDoc = wrdApp.Documents.Open(fileName:=tmpFileName, Visible:=True)

'use T11 to distiguish between affected and non affected. Affected print at A3 landscape
T11_EXCEPTION:
    If T11Affected = True Then
            With wrdApp.ActiveDocument.PageSetup
                .PaperSize = wdPaperA3
                .Orientation = wdOrientLandscape
            End With
    End If

With wrdApp.ActiveDocument.PageSetup
    .TopMargin = wrdApp.InchesToPoints(0.5)
    .BottomMargin = wrdApp.InchesToPoints(0.5)
    .LeftMargin = wrdApp.InchesToPoints(0.5)
    .RightMargin = wrdApp.InchesToPoints(0.5)
End With


'if email has attachements that need to be saved, add the sName to the if statement below
'    If sName = "T03" Or sName = "T24" Or sName = "H05" Or sName = "H35" Then
'        strToSaveAs = myDocs & "\" & sName & ".01.pdf" 'check for duplicate filenames
'        strToSaveAs2 = myDocs & "\" & sName & ".pdf"
' if matched, add the current time to the file name
'    Else
'        strToSaveAs = myDocs & "\" & sName & ".pdf"
'        strToSaveAs2 = myDocs & "\" & sName & ".01.pdf"
'        Debug.Print "StrFile2 = " & strToSaveAs2
'    End If
'

strToSaveAs = myDocs & "\" & sName & ".00.pdf" 'check for duplicate filenames
strToSaveAs2 = myDocs & "\" & sName & ".pdf"


If fso.FileExists(strToSaveAs2) Then
    Debug.Print "File already exists!"
    fso.DeleteFile strToSaveAs2, True
    Debug.Print "Deleting duplicate...."
ElseIf fso.FileExists(strToSaveAs) Then
    Debug.Print "File already exists!"
    fso.DeleteFile strToSaveAs, True
    Debug.Print "Deleting duplicate...."
End If



'  If MsgBox("Do you want to overide file?", vbQuestion + vbYesNo, "Overide") = vbNo Then
'    sName = "Exists!"
'    GoTo CLOSE_WORD
'    End If
    


wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
strToSaveAs, ExportFormat:=wdExportFormatPDF, _
OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
Range:=wdExportAllDocument, From:=0, To:=0, Item:= _
wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
BitmapMissingFonts:=True, UseISO19005_1:=False

CLOSE_WORD:
wrdDoc.Close
wrdApp.Quit
Set wrdDoc = Nothing
Set wrdApp = Nothing

End Function


' This function removes invalid and other characters from file names
Private Sub ReplaceCharsForFileName(sName As String, sChr As String)
sName = Replace(sName, "/", sChr)
sName = Replace(sName, "\", sChr)
sName = Replace(sName, ":", sChr)
sName = Replace(sName, "?", sChr)
sName = Replace(sName, Chr(34), sChr)
sName = Replace(sName, "<", sChr)
sName = Replace(sName, ">", sChr)
sName = Replace(sName, "|", sChr)
sName = Replace(sName, "&", sChr)
sName = Replace(sName, "%", sChr)
sName = Replace(sName, "*", sChr)
sName = Replace(sName, " ", sChr)
sName = Replace(sName, "{", sChr)
sName = Replace(sName, "[", sChr)
sName = Replace(sName, "]", sChr)
sName = Replace(sName, "}", sChr)
sName = Replace(sName, "!", sChr)

End Sub
