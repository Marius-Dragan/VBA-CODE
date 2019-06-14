Attribute VB_Name = "ConsolidatingWS"
Option Explicit
Sub ConsolidatingMultipleWS()

Dim MasterWS As Worksheet
Dim ws As Worksheet
Dim CopyRange As Variant
Dim lrowOtherSheets As Long
Dim LRow As Long

Set MasterWS = Sheets(1)
Set ws = ActiveSheet

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> MasterWS.Name Then
                ws.Activate
                 Application.ScreenUpdating = False
                 
                With ws
                    lrowOtherSheets = ws.Range("C" & ws.Rows.Count).End(xlUp).Row
                End With
                 
                    'Set CopyRange = WS.Range("A2:N" & lrowOtherSheets).Copy
                    Set CopyRange = ws.UsedRange
                        CopyRange.Copy
                    
                    With MasterWS
                            If LRow = 0 Then
                                LRow = .Range("A" & .Rows.Count).End(xlUp).Row
                            Else
                                LRow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
                            End If
                    End With
                    
                    MasterWS.Range("A" & LRow).PasteSpecial Paste:=xlPasteAll, Transpose:=False
                    Application.CutCopyMode = False
                    
            End If
                 Application.ScreenUpdating = True
        Next ws
    
    MsgBox "Process completed!", vbInformation
End Sub
