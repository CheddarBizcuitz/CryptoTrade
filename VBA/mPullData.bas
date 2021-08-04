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
  
    Dim ws As Worksheet, Sh As Worksheet
    Dim Rng As Range, c As Range

    Set ws = Sheets("Master")
    Set Rng = ws.Range("B7:B100")
    
    For Each Sh In Sheets
        For Each c In Rng.Cells
            If Sh.Name = c Then
                c.Offset(0, 1) = Sh.Range("AC44").Value
                c.Offset(0, 2) = Sh.Range("S3").Value
                c.Offset(0, 3) = Sh.Range("T3").Value
                c.Offset(0, 4) = Sh.Range("U3").Value
                c.Offset(0, 5) = Sh.Range("V3").Value
                c.Offset(0, 6) = Sh.Range("W3").Value
                c.Offset(0, 11) = Sh.Range("S17").Value
                c.Offset(0, 12) = Sh.Range("S18").Value
                c.Offset(0, 13) = Sh.Range("AC24").Value
            End If
        Next c
    Next Sh
     
    ActiveSheet.Range("Table1").Borders.LineStyle = xlNone
    
    Range("B7").Select
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
