VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTransactionCheck 
   Caption         =   "Transaction Check"
   ClientHeight    =   3144
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4284
   OleObjectBlob   =   "ufTransactionCheck.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufTransactionCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()

    ListBox1.RowSource = "MethodListString"

End Sub
