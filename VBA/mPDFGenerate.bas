Attribute VB_Name = "mPDFGenerate"
Option Explicit

Sub PDFGenerate()

Dim PDFTempFile, PDFSaveFolder, PDFNewName As String
Dim DataRow, LastRow As Long

    With ActiveSheet
        LastRow = Range("B9999").End(xlUp).Row
        PDFTempFile = Range("F3").Value
        PDFSaveFolder = Range("F6").Value
        
        ThisWorkbook.FollowHyperlink "" & PDFTempFile & ""
        DoEvents
        
        For DataRow = 3 To 3 'LastRow
        
        Application.SendKeys "{Tab}", True
        Application.SendKeys .Range("B3").Value, True
            
            
            
            
            
            
            Next DataRow
        
        
        





     End With

End Sub
