VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufDeleteSheet 
   Caption         =   "Delete Worksheet?"
   ClientHeight    =   2040
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ufDeleteSheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufDeleteSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fConfirm_Click()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.DisplayAlerts = False

    If Len(Me.fCoin.Value) = 0 Then
        MsgBox "You must input a value to proceed.", vbOKOnly, "Hello."
        Me.fCoin.SetFocus
            Exit Sub
    End If

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        If Me.fCoin.Value = ws.Name Then

            ws.Delete

        End If

            Next ws
     
    Me.fCoin.Value = ""

    Me.Hide

Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.DisplayAlerts = True

End Sub
