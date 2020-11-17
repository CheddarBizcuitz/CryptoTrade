VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()

    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
End Sub
Private Sub bAddRow_Click()

ActiveSheet.Unprotect

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    Rows(Range("TempTableData").Rows.count + Range("TempTableData").Cells(1, 1).Row - 1).Select
    Range("B" & ActiveCell.Row & " :P" & ActiveCell.Row).Select
        Selection.ListObject.ListRows.Add AlwaysInsert:=True
    
    Range("H4:P4").Select
        Selection.AutoFill Destination:=Range("FormulaFill")

    Range("TempTableData").Select
        Selection.RowHeight = 18
    
    Range("B4").Select
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    
End Sub

Private Sub bManList_Click()

    ufListings.Show

End Sub

Private Sub bRemoveRow_Click()

ActiveSheet.Unprotect

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
Application.DisplayAlerts = False
    
    Dim MyRg As Range
    Set MyRg = Range("TempTableData")
    
    If Intersect(MyRg, ActiveCell) Is Nothing Or Not Intersect(ActiveCell, Range("B4:G5")) Is Nothing Then
        MsgBox "Unable to delete cell.", vbOKOnly, "Hello."
        
     Else
 
        Range("B" & ActiveCell.Row & " :P" & ActiveCell.Row).Delete
 
     On Error Resume Next
 
    End If
     
    Range("B4").Select

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
Application.DisplayAlerts = True

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub

Private Sub bRefreshPrice_Click()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False

    ActiveWorkbook.RefreshAll

Application.ScreenUpdating = True
Application.DisplayStatusBar = True

End Sub
