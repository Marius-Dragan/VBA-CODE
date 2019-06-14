Attribute VB_Name = "M3APIForDO"
Option Explicit
Sub ConsolidatingMultipleWS_M3_DO()

Dim MasterWS As Worksheet
Dim ws As Worksheet
Dim CopyRange As Variant
Dim lrowOtherSheets As Long
Dim LRow As Long

Set MasterWS = ActiveSheet
Set ws = ActiveSheet

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> MasterWS.Name Then
                ws.Activate
                ws.Tab.Color = vbGreen
                 Application.ScreenUpdating = False
                 
                With ws
                    lrowOtherSheets = ws.Range("C" & ws.Rows.Count).End(xlUp).Row
                End With
                 
                    'Set CopyRange = WS.Range("A2:N" & lrowOtherSheets).Copy
                    Set CopyRange = ws.UsedRange.Offset(1, 0)
                        CopyRange.Copy
                    
                    With MasterWS
                            LRow = .Range("A" & .Rows.Count).End(xlUp).Row
                            If LRow = 0 Then
                                LRow = .Range("A" & .Rows.Count).End(xlUp).Row
                            Else
                                LRow = .Range("A" & .Rows.Count).End(xlUp).Row + 1
                            End If
                    End With
                    'MasterWS.Activate
                    MasterWS.Range("A" & LRow).PasteSpecial Paste:=xlPasteAll, Transpose:=False
                    Application.CutCopyMode = False
                    
            End If
        Next ws
        
        MasterWS.Name = "MASTER_WS"
        MasterWS.Activate
        MasterWS.Range("N2:N" & LRow).Formula = "=CONCATENATE(""000000"",A2)"
        MasterWS.Range("A2:A" & LRow).NumberFormat = "@"
        MasterWS.Range("N2:N" & LRow).NumberFormat = "@"
        MasterWS.Range("N2:N" & LRow).value = MasterWS.Range("N2:N" & LRow).value
        Set CopyRange = MasterWS.Range("N2:N" & LRow)
            CopyRange.Copy
        MasterWS.Range("A2").PasteSpecial Paste:=xlPasteAll, Transpose:=False
        Application.CutCopyMode = False
        MasterWS.Range("N2:N" & LRow).ClearContents
        
        Application.ScreenUpdating = True
        MasterWS.Range("A1").Activate
        MsgBox "Process completed!", vbInformation
End Sub
Sub CopyDataToTemplate_M3_API_DO()

Dim ws As Worksheet
Dim srcWB As Workbook
Dim destWB As Workbook
Dim srcWS As Worksheet
Dim destWS As Worksheet
Dim CopyRange As Variant

Dim i As Long, j As Long
Dim srcLRow As Long, destLRow As Long

Set destWB = Excel.Workbooks("DEV_Template_OD_MMS100MI_AddDOLine_example.xlsx")
Set srcWB = ActiveWorkbook
Set srcWS = srcWB.ActiveSheet
Set destWS = destWB.Sheets("MMS100MIAddDOLine")

srcLRow = srcWS.Cells(srcWS.Rows.Count, "A").End(xlUp).Row
destLRow = destWS.Cells(destWS.Rows.Count, "A").End(xlUp).Row

destWS.Range("H3").value = "WHSL1"
destWS.Range("I3").value = "WHSL2"
Application.ScreenUpdating = False
'loop through column 1 to 19
For i = 1 To 19
    For j = 1 To 13
        'loop through columns
        
            If destWS.Cells(3, i).value = srcWS.Cells(1, j).value Then
            'Debug.Print destWS.Cells(3, i).value
            'Debug.Print srcWS.Cells(1, j).value
                ' Copy column B to Column D as written in your code above
                Set CopyRange = srcWS.Range(Cells(2, j), Cells(srcLRow, j))
                    CopyRange.Copy
                ' paste columns from one wb to Columns to another wb
                destWS.Cells(destLRow, i).PasteSpecial Paste:=xlPasteAll, Transpose:=False
                Application.CutCopyMode = False
            End If
            'Debug.Print destWS.Cells(3, i).value
            'Debug.Print srcWS.Cells(1, j).value
    Next j
Next i
destWS.Range("H3").value = "WHSL"
destWS.Range("I3").value = "TWSL"
Application.ScreenUpdating = True
MsgBox "Process completed!", vbInformation
End Sub
