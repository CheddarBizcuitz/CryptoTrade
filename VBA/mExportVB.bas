Attribute VB_Name = "mExportVB"
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
' Requires enabling "Microsoft Scripting Runtime" in Visual Basic/Tools/References
Option Explicit

Public Sub ExportVisualBasicCode()

    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
        Dim WshShell As Object
        Dim fsos As Object
        Dim SpecialPath As String

        Set WshShell = CreateObject("WScript.Shell")
        Set fsos = CreateObject("scripting.filesystemobject")

        SpecialPath = WshShell.SpecialFolders("C:\")
    
        directory = SpecialPath & "\VisualBasic_Export"
        count = 0
        
    If ActiveWorkbook.Name = "Master.xlsm" Then
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
                           
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeValue("00:00:05"), "clearStatusBar"
    
    Call ClearStatusBar
    
     Else
     
        MsgBox "Cannot export files, workbook name must be 'Master' to continue.", vbOKOnly, "Caution!"
        
    End If

End Sub

Sub ClearStatusBar()

    Application.StatusBar = False

End Sub
