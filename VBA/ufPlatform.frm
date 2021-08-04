VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPlatform 
   Caption         =   "Platform List"
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3984
   OleObjectBlob   =   "ufPlatform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPlatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pBack_Click()

    Me.Hide
    ufListings.Show

End Sub

Private Sub UserForm_Activate()

    Me.ListBox1.RowSource = "PlatList"

End Sub

Private Sub pAdd_Click()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Listings.Visible = True

Listings.Activate

    If Me.TextBox1.Value = "" Then
        MsgBox "You must add a value to continue."

     Else

    Range("B3:B3").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B3").Value = Me.TextBox1.Value
    Range("B3").Select
        Selection.HorizontalAlignment = xlCenter
        Selection.Locked = False
   
       ActiveWorkbook.Worksheets("DataEntryList").ListObjects("tPlatform").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("DataEntryList").ListObjects("tPlatform").Sort. _
        SortFields.Add2 Key:=Range("tPlatform[[#All],[Platform]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DataEntryList").ListObjects("tPlatform").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                
    End If
      
    Me.TextBox1.Value = ""
      
    Me.Hide
    
Listings.Visible = False
      
    Me.Show
    
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
      
End Sub

Private Sub pRemove_Click()

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

Listings.Visible = True

Listings.Activate

 Listings.Range("tPlatform").Select
    Dim lastrows As Long
    lastrows = Listings.Cells(Rows.count, 1).End(xlUp).Row
    Dim strRange As String
        With ListBox1
    strRange = .RowSource
    Range(strRange).Cells(.ListIndex + 1, 1).Delete Shift:=xlUp
        .RowSource = vbNullString
        .RowSource = strRange
        
    End With
 
    Me.Hide
    
Listings.Visible = False
      
    Me.Show

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
    
End Sub

Private Sub pDeselect_Click()

    ListBox1.Selected(ListBox1.ListIndex) = False

End Sub
