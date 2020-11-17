Attribute VB_Name = "mDeleteCoin"
Option Explicit

Sub DeleteCoin()

ActiveSheet.Unprotect

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.DisplayAlerts = False
    
    Dim MyRg As Range
    Set MyRg = Range("Table1")
 
    If Intersect(MyRg, ActiveCell) Is Nothing Then
        MsgBox "Unable to delete cell.", vbOKOnly, "Hello."
     Else
 
    Range("B" & ActiveCell.Row & " :H" & ActiveCell.Row).Delete
 
     On Error Resume Next
 
    End If
     
    CoinList.Range("B7").Select

Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.DisplayAlerts = True

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
