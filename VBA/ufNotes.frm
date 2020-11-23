VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNotes 
   ClientHeight    =   8412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8544
   OleObjectBlob   =   "ufNotes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()

    ActiveWorkbook.FollowHyperlink Address:="https://www.irs.gov/pub/irs-access/f8949_accessible.pdf", NewWindow:=True
    Unload Me

End Sub

Private Sub Label4_Click()

    ActiveWorkbook.FollowHyperlink Address:="https://www.irs.gov/pub/irs-pdf/f1040sd.pdf", NewWindow:=True
    Unload Me

End Sub
