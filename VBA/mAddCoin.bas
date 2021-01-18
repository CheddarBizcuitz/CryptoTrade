Attribute VB_Name = "mAddCoin"
Option Explicit

Sub AddCoin()

ActiveSheet.Unprotect

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    If CoinList.ComboBox1.Text <> "" And CoinList.Range("B7").Value <> "" Then
        MsgBox "Please add blank cell before data re-entry. See 'Notes' for more details.", vbOKOnly, "Hello."
     Else

    If IsNumeric(Application.Match(CoinLibrary.Range("B3"), CoinLibrary.Range("B4:B150"), 0)) Then
        CoinList.Range("B7:H7").Select
            Selection.Locked = False
        CoinList.Range("B7").Value = CoinList.ComboBox1.Value
            Selection.Locked = True
        CoinList.Range("B7").Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Locked = False
     Else
     
        MsgBox "Invalid input. Generating blank cell if required.", vbOKOnly, "Hello."
     
    End If

    CoinList.ComboBox1.Value = ""
    CoinList.Range("Q2").Value = ""

    If CoinList.Range("B7") <> "" Then
        CoinList.Range("B7:H7").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
   
     End If
     
    CoinList.Range("B7").Select
     
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    
End Sub
