Attribute VB_Name = "mPDFSaveFolder"
Option Explicit

Sub PDFSaveFolder()

Dim PDFFldr As FileDialog
Set PDFFldr = Application.FileDialog(msoFileDialogFolderPicker)

    With PDFFldr
        .Title = "Select Save Location"
    
    If .Show <> -1 Then GoTo NoSelection
        AutofillPDF.Range("F6").Value = .SelectedItems(1)
        
     End With
     
NoSelection:

End Sub
