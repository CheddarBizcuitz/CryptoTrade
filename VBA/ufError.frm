VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufError 
   Caption         =   "Error"
   ClientHeight    =   3552
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3828
   OleObjectBlob   =   "ufError.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pBack_Click()

    Me.Hide
    ufListings.Show

End Sub

Private Sub pHelp_Click()

    MsgBox "Enter your account's 'actual' coin volume to determine the error within this sheet. Please note that any value other than 0 will overwrite the sheet's coin balance value.", vbOKOnly, "Hello."
 
End Sub

Private Sub UserForm_Activate()

    Me.eLabelFront1.Caption = ActiveSheet.Range("AC28")
    Me.eLabelFront2.Caption = ActiveSheet.Range("AC34")
    Me.eLabelFront3.Caption = ActiveSheet.Range("AC38")
    Me.eLabelFront4.Caption = ActiveSheet.Range("AC42")

End Sub

Private Sub pAlter_Click()

ActiveSheet.Unprotect

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    If Me.TextBox1.Value = "" Then
        MsgBox "You must input a value to continue."
     Else

    Range("Error").Select
        Dim lastrows As Long
            lastrows = Listings.Cells(Rows.count, 1).End(xlUp).Row
        Dim strRange As String
            With eLabelFront1
                strRange = .Caption
                .Caption = vbNullString
                .Caption = strRange
 
    Selection = Me.TextBox1.Value
  
    Me.TextBox1.Value = ""
 
            End With
 
    End If
    
    Me.Hide
    Me.Show
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
 
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
