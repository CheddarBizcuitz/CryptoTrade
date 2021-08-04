VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTransaction 
   Caption         =   "Transaction Type List"
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3984
   OleObjectBlob   =   "ufTransaction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pBack_Click()

    Me.Hide
    ufListings.Show

End Sub

Private Sub pCheck_Click()

    ufTransactionCheck.Show

End Sub

Private Sub pHelp_Click()

    ufTransactionHelp.Show

End Sub

Private Sub UserForm_Activate()

    Me.ListBox1.RowSource = "MethodList"

End Sub

Private Sub pAlter_Click()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Listings.Visible = True

Listings.Activate
 
    If Me.TextBox1.Value = "" Then
        MsgBox "You must input a value to continue."
     Else

    Range("MethodList").Select
        Dim lastrows As Long
            lastrows = Listings.Cells(Rows.count, 1).End(xlUp).Row
        Dim strRange As String
         With ListBox1
            strRange = .RowSource
            Range(strRange).Cells(.ListIndex + 1, 1).Select
            .RowSource = vbNullString
            .RowSource = strRange
 
    Selection = Me.TextBox1.Value
  
    Me.TextBox1.Value = ""
 
            End With
 
    End If

    Me.Hide
    
Listings.Visible = False
      
    Me.Show

Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub
