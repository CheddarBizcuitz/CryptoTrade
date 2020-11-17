VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufListings 
   Caption         =   "Manage Worksheet"
   ClientHeight    =   2652
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ufListings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufListings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub addmultirow()

ActiveSheet.Unprotect

    Rows(Range("TempTableData").Rows.count + Range("TempTableData").Cells(1, 1).Row - 1).Select
    Range("B" & ActiveCell.Row & " :P" & ActiveCell.Row).Select
        Selection.ListObject.ListRows.Add AlwaysInsert:=True
    
    Range("H4:P4").Select
        Selection.AutoFill Destination:=Range("FormulaFill")

    Range("TempTableData").Select
        Selection.RowHeight = 18
    
    Range("B4").Select
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    
End Sub

Private Sub clkAddMultiRow_Click()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    Dim I As Long
    Dim N As Variant

        N = InputBox("How many rows?")

        If Not IsNumeric(N) Then Exit Sub

        For I = 1 To N
            addmultirow
        Next I
        
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

End Sub

Private Sub clkErrorMargin_Click()

    Me.Hide
    ufError.Show

End Sub

Private Sub clkPlat_Click()

    Me.Hide
    ufPlatform.Show

End Sub

Private Sub clkTran_Click()

    Me.Hide
    ufTransaction.Show

End Sub


Private Sub clkInvest_Click()

    Me.Hide
    ufInvestment.Show

End Sub

Private Sub clkSB_Click()

    Me.Hide
    ufSmallBal.Show

End Sub
