Attribute VB_Name = "mPullData"
Option Explicit

Sub PullData()

ActiveSheet.Unprotect

ActiveWorkbook.RefreshAll
DoEvents

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
  
    Dim ws As Worksheet, sh As Worksheet
    Dim Rng As Range, c As Range

    Set ws = Sheets("Master")
    Set Rng = ws.Range("B7:B100")
    
    For Each sh In Sheets
        For Each c In Rng.Cells
            If sh.Name = c Then
                c.Offset(0, 1) = sh.Range("R3").Value
                c.Offset(0, 2) = sh.Range("S3").Value
                c.Offset(0, 3) = sh.Range("T3").Value
                c.Offset(0, 4) = sh.Range("U3").Value
                c.Offset(0, 5) = sh.Range("V3").Value
                c.Offset(0, 6) = sh.Range("W3").Value
                c.Offset(0, 11) = sh.Range("S17").Value
                c.Offset(0, 12) = sh.Range("S18").Value
                c.Offset(0, 13) = sh.Range("AC24").Value
            End If
        Next c
    Next sh
     
    ActiveSheet.Range("Table1").Borders.LineStyle = xlNone
    
    ActiveSheet.Range("Table1[Investment]").Select
    Selection.Locked = True
    Selection.FormulaHidden = True
    
    Range("B7").Select
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
