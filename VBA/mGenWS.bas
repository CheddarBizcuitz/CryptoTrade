Attribute VB_Name = "mGenWS"
' credit: https://www.mrexcel.com/board/threads/vba-code-to-create-a-new-sheet-from-a-template-and-rename-it-from-a-list.813294/
Option Explicit

Sub GenWS()

ActiveSheet.Unprotect

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False

    Dim wsMASTER As Worksheet, wsTEMP As Worksheet, wasVISIBLE As Boolean
    Dim shNAMES As Range, Nm As Range

    With ThisWorkbook                                               'keep focus in this workbook
        Set wsTEMP = .Sheets("Temp")                                'sheet to be copied
        wasVISIBLE = (wsTEMP.Visible = xlSheetVisible)              'check if it's hidden or not
        If Not wasVISIBLE Then wsTEMP.Visible = xlSheetVisible      'make it visible
    
        Set wsMASTER = .Sheets("Master")                            'sheet with names
                                                                'range to find names to be checked
        Set shNAMES = wsMASTER.Range("B7:B" & Rows.count).SpecialCells(xlConstants)     'or xlFormulas
    
        Application.ScreenUpdating = False                              'speed up macro
        For Each Nm In shNAMES                                          'check one name at a time
            If Not Evaluate("ISREF('" & CStr(Nm.Text) & "'!A1)") Then   'if sheet does not exist...
                wsTEMP.Copy After:=.Sheets(.Sheets.count)               '...create it from template
                ActiveSheet.Name = CStr(Nm.Text)                        '...rename it
            End If
             Next Nm
    
    wsMASTER.Activate                                           'return to the master sheet
    If Not wasVISIBLE Then wsTEMP.Visible = xlSheetHidden       'hide the template if necessary
    Application.ScreenUpdating = True                           'update screen one time at the end
 
     End With

    Dim R As Range, rowz As Long, I As Long
    Set R = ActiveSheet.Range("Table1")
    rowz = R.Rows.count
    For I = rowz To 1 Step (-1)
        If WorksheetFunction.CountA(R.Rows(I)) = 0 Then R.Rows(I).Delete
     Next

    MsgBox "All sheets created"
    
    Range("B7").Select

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True

ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

End Sub
