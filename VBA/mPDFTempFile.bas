Attribute VB_Name = "mPDFTempFile"
Option Explicit

Sub PDFTempFile()

Dim PDFFldr As FileDialog
Set PDFFldr = Application.FileDialog(msoFileDialogFilePicker)

    With PDFFldr
        .Title = "Select Template File"
        .Filters.Add "PDF Type Files", "*.pdf", 1
    
    If .Show <> -1 Then GoTo NoSelection
        AutofillPDF.Range("F3").Value = .SelectedItems(1)
    
     End With

NoSelection:

End Sub
