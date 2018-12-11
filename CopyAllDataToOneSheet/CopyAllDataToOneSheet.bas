Attribute VB_Name = "CopyAllDataToOneSheet"
Option Explicit
 
Sub CopyAllDataToOneSheet()
 
Dim wsAllStaff As Worksheet
Dim WS As Worksheet
Dim lrowOtherSheets As Long
Dim lrow As Long
 
Set wsAllStaff = Sheets("AllStaff")
Set WS = ActiveSheet
 
    For Each WS In ActiveWorkbook.Worksheets
        If WS.Name <> wsAllStaff.Name Then
                WS.Activate
               
                 Application.ScreenUpdating = False
                
                With WS
                    lrowOtherSheets = WS.Range("A" & WS.Rows.Count).End(xlUp).Row
                End With
              
                 
                    WS.Range("A10:N" & lrowOtherSheets).Copy
                   
                     With WS
                            lrow = wsAllStaff.Range("A" & wsAllStaff.Rows.Count).End(xlUp).Row + 1
                    End With
                   
                    wsAllStaff.Range("A" & lrow).PasteSpecial Paste:=xlPasteAll, Transpose:=False
                    Application.CutCopyMode = False
                   
                 
            End If
                 Application.ScreenUpdating = True
        Next WS
 
End Sub
