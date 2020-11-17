VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCustomCoin 
   Caption         =   "New Currency - Input Form"
   ClientHeight    =   3384
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4788
   OleObjectBlob   =   "ufCustomCoin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufCustomCoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fConfirm_Click()

CoinList.Unprotect

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    If Len(Me.fCoin.Value) = 0 Then
        MsgBox "You must input a value to continue.", vbOKOnly, "Hello."
     Else

    If Me.fCoin.Value <> "" And CoinList.Range("B7").Value <> "" Then
        MsgBox "Please add blank row before data re-entry. See 'Notes' for more details.", vbOKOnly, "Hello."
     Else

    If Me.fSymbol.Value = "" Then
        Range("B7").Value = Me.fCoin.Value
        Range("B7").Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Locked = False
        CoinList.Range("B7:H7").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
     Else

        Range("B7").Value = Me.fCoin.Value & " " & "(" & Me.fSymbol.Value & ")"
        Range("B7").Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Locked = False
        CoinList.Range("B7:H7").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
     
     End If
     
      End If
      
    Me.fCoin.Value = ""
    Me.fSymbol.Value = ""
    
    Me.Hide
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

CoinList.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
