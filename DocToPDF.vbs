VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SaveAllAsPDF()

Dim strFilename As String
 Dim strDocName As String
 Dim strPath As String
 Dim oDoc As Document
 Dim fDialog As FileDialog
 Dim intPos As Integer
 Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
 With fDialog
     .Title = "Select folder and click OK"
     .AllowMultiSelect = False
     .InitialView = msoFileDialogViewList
     If .Show <> -1 Then
         MsgBox "Cancelled By User", , "List Folder Contents"
         Exit Sub
     End If
     strPath = fDialog.SelectedItems.Item(1)
     If Right(strPath, 1) <> "\" Then strPath = strPath + "\"
 End With
 If Documents.Count > 0 Then
     Documents.Close SaveChanges:=wdPromptToSaveChanges
 End If
 If Left(strPath, 1) = Chr(34) Then
     strPath = Mid(strPath, 2, Len(strPath) - 2)
 End If
 strFilename = Dir$(strPath & "*.doc*")
 While Len(strFilename) <> 0
     Set oDoc = Documents.Open(strPath & strFilename)
     strDocName = ActiveDocument.FullName
     intPos = InStrRev(strDocName, ".")
     strDocName = Left(strDocName, intPos - 1)
'This instruction converts to PDF

       strDocName = strDocName & ".pdf"
     oDoc.SaveAs FileName:=strDocName, _
         FileFormat:=wdFormatPDF
         
     oDoc.Close SaveChanges:=wdDoNotSaveChanges
     strFilename = Dir$()
 Wend
 End Sub
 

