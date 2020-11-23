VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGoToWS 
   Caption         =   "Worksheets"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4500
   OleObjectBlob   =   "ufGoToWS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufGoToWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pGoTo_Click()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    Dim ws As Worksheet
    Dim wsname As String

    wsname = ListBox1.Value

    If ListBox1.Selected(ListBox1.ListIndex) = True Then
    
        For Each ws In Application.Worksheets
    
            If wsname = ws.Name Then
            
                ws.Activate
            
            End If
            
         Next ws
    
    End If
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub

Private Sub UserForm_Activate()

    Dim ws As Worksheet
    
    For Each ws In Application.Worksheets

        If ws.Visible = True Then
            ListBox1.AddItem ws.Name
    
         End If
    
     Next ws

End Sub

Private Sub UserForm_Terminate()

    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
End Sub
