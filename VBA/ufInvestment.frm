VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufInvestment 
   Caption         =   "Investment List"
   ClientHeight    =   3420
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "ufInvestment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufInvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pAdd_Click()

ActiveSheet.Unprotect

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    If Me.TextBox1.Value <> "" And Me.TextBox2.Value <> "" Then
 
        ActiveSheet.Range("Y6:AA6").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
 
        Range("Y6").Value = Me.TextBox1.Value
        Range("Y6").Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Locked = False
        Range("Z6").Value = Me.TextBox2.Value
        Range("Z6").Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Locked = False
                    
     Else
        MsgBox "You must add a value to continue."
                
    End If
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
      
    Me.TextBox1.Value = ""
    Me.TextBox2.Value = ""
      
    Me.Hide
    Me.Show
      
ActiveSheet.Protect

End Sub

Private Sub pBack_Click()

    Me.Hide
    ufListings.Show

End Sub

Private Sub pRemove_Click()

ActiveSheet.Unprotect

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
    
    ActiveSheet.Range("tInvestmentsData").Select
        Dim lastrows As Long
            lastrows = ActiveSheet.Cells(Rows.count, 1).End(xlUp).Row
        Dim strRange As String
            With ListBox1
        strRange = .RowSource
        Range(strRange).Cells(.ListIndex + 1, 1).Delete Shift:=xlUp
        .RowSource = vbNullString
        .RowSource = strRange
 
     End With
     
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
 
ActiveSheet.Protect

End Sub

Private Sub UserForm_Activate()

    Me.ListBox1.RowSource = "InvestList"

End Sub
