Attribute VB_Name = "mImportVB"
' Primary Source - https://www.rondebruin.nl/win/s9/win002.htm
' Requires enabling "Microsoft Visual Basic For Applications Extensibility 5.3" and "Microsoft Scripting Runtime" in Visual Basic/Tools/References
Option Explicit

Public Sub ImportModules()

    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If
    
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
        
    Next objFile
    
End Sub

Function FolderWithVBAProjectFiles() As String

    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("C:\")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If fso.FolderExists(SpecialPath & "VisualBasic") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VisualBasic"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "VisualBasic") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VisualBasic"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()

        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
        
End Function
