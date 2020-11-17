VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoinList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub ComboBox1_GotFocus()
    
    ComboBox1.ListFillRange = "DropDownList"
    Me.ComboBox1.DropDown
    
End Sub

Private Sub Worksheet_()
    
    If Selection = CoinList.ComboBox1.Select Then
        Me.ComboBox1.DropDown
    
    End If
    
End Sub

Private Sub HelpButton_Click()
    
    ufNotes.Show
    
End Sub

Private Sub Worksheet_Activate()

Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.DisplayAlerts = False
      
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
    
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.DisplayAlerts = True

End Sub

Private Sub Worksheet_Deactivate()

    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
End Sub
