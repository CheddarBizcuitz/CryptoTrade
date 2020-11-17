VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Listings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_()

    Dim R As Range, rowz As Long, I As Long
        Set R = Listings.Range("tPlatform")
        rowz = R.Rows.count
    For I = rowz To 1 Step (-1)
        If WorksheetFunction.CountA(R.Rows(I)) = 0 Then R.Rows(I).Delete
     Next

End Sub

