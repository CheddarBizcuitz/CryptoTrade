VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSmallBal 
   Caption         =   "Small Balance"
   ClientHeight    =   2424
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3984
   OleObjectBlob   =   "ufSmallBal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSmallBal"
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

    MsgBox "Enter 'Small Balance' to set the coin's account balance limit. Small balances can range from .0001 to 100+. Small balances can differ based on the coin's volume and market cap.", vbOKOnly, "Hello."

End Sub

Private Sub UserForm_Activate()

    Me.ListBox1.RowSource = "SmallBalance"
    Me.ListBox1.ListIndex = -1

End Sub

Private Sub pAlter_Click()

ActiveSheet.Unprotect

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    If Me.TextBox1.Value = "" Or ListBox1.ListIndex = -1 Then
        MsgBox "You must input a value to continue."
     Else

    Range("SmallBalance").Select
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
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
 
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
