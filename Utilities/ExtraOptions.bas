Attribute VB_Name = "ExtraOptions"
Option Explicit
Sub FillBlanks()
Dim WS As Worksheet
Dim ColumnLetter As String
Dim fillingValue As String
Dim LRow As Long

Set WS = ActiveSheet
ColumnLetter = UCase(Application.InputBox("Input the column letter where you want the data to be filled:", "Column Letter", Type:=2))
If IsNumeric(ColumnLetter) = True Or ColumnLetter = "" Then
    MsgBox "Not a letter or blank. Please try again!", vbInformation, "Fill Blanks"
    Exit Sub
ElseIf ColumnLetter = "FALSE" Then
    Exit Sub
End If
fillingValue = Application.InputBox("Input the vaule for the blank cells to be filled with:", "Fill blank cells", Type:=2)
If fillingValue = "" Or fillingValue = "False" Then
    MsgBox "No data in the input box. Please try again!", vbInformation, "Fill Blanks"
    Exit Sub
End If
Call FindLastRowCol(WS, LRow, 1)

If LRow > 0 Then
    With WS.Range(ColumnLetter & 1, ColumnLetter & LRow)
        On Error GoTo ErrorHandler
        If .Application.CountBlank(.Cells) > 0 Then
            .SpecialCells(xlCellTypeBlanks).value = fillingValue
        End If
    End With
Else
    MsgBox "Error: Could not find last row! Your sheet might be blank.", vbCritical, "Fill Blanks"
End If
Exit Sub
ErrorHandler:
Select Case Err.Number
Case 1004
    MsgBox "No cells were found! Please make sure there is at least 1 cell that has data in the column.", vbInformation, "Fill Blanks"
End Select
Exit Sub
End Sub
Sub ClearAllLogs()
Dim LRow As Long
    With ActiveSheet
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 16 Then
            LRow = LRow + 1
        End If
        .Range("A17:B" & LRow).ClearContents
    End With
End Sub

Sub ClearNokLogsOnly()
Dim ClearRng As Range
Dim LRow As Long
Dim i As Long
With ActiveSheet
    LRow = .Range("A" & .Rows.Count).End(xlUp).Row
    If LRow = 17 Then
            LRow = LRow + 1
    End If
    For i = 17 To LRow
        If UCase(.Range("A" & i).value) = "NOK" Then
            If ClearRng Is Nothing Then
                Set ClearRng = .Range(.Cells(i, 1), .Cells(i, 2))
            Else
                Set ClearRng = Union(ClearRng, .Range(.Cells(i, 1), .Cells(i, 2)))
            End If
        End If
    Next i
    If Not ClearRng Is Nothing Then ClearRng.Clear
    Set ClearRng = Nothing
End With
End Sub

Sub ClearData()
Dim WS As Worksheet
Dim TotalColumns As Long
Dim xTotalColumns As Long
Dim LRow As Long
Dim Rng As Range
Dim questionBoxPopUp As VbMsgBoxResult
Set WS = ActiveSheet

questionBoxPopUp = MsgBox("You are attempting to clear the data from worksheet. After the data is cleared there is no undo action that can be done. Are you sure you want to continue?", vbQuestion + vbYesNo + vbDefaultButton1, "Clear worksheet data")
If questionBoxPopUp = vbNo Then Exit Sub


TotalColumns = WS.Cells(16, Columns.Count).End(xlToLeft).Column
xTotalColumns = Cells.SpecialCells(xlCellTypeLastCell).Column
If TotalColumns < 3 Or TotalColumns < xTotalColumns Then TotalColumns = xTotalColumns

LRow = FindLastRow(WS, 3, TotalColumns)
If LRow < 17 Then LRow = 17

Set Rng = WS.Range(Cells(17, 3), Cells(LRow, TotalColumns))
If LRow > 16 Then
    Rng.ClearContents
    Rng.ClearFormats
End If
End Sub
Sub UpdateEndRow()
    With ActiveSheet
        .Range("B8").value = FindLastRow(ActiveSheet, 3, .UsedRange.Columns.Count)
    End With
End Sub
Function UploadStatus(ByRef WS As Worksheet, ByVal StartRow As Long, ByVal EndRow As Long, Optional ByVal strMsg As String) As String
 Dim okCount As Long
 Dim nokCount As Long

 If StartRow = 0 And EndRow = 0 Then
   Exit Function
 End If
 
 With WS.Application.WorksheetFunction
    okCount = .CountIf(Range("A" & StartRow & ":A" & EndRow), "OK")
    nokCount = .CountIf(Range("A" & StartRow & ":A" & EndRow), "NOK")
    UploadStatus = .TextJoin(" ", True, strMsg, okCount, "OK", "row" & IIf(okCount > 1, "s", ""), ",", nokCount, "NOK", "row" & IIf(nokCount > 1, "s", ""))
 End With
End Function
Function FindLastRow(ByVal WS As Worksheet, Optional ByVal FromCol As Long = 0, Optional ByVal ToCol As Long = 0) As Long
Dim i As Long
Dim lastRow As Long
If FromCol = 0 Then FromCol = 3
If ToCol = 0 Then ToCol = 10
For i = FromCol To ToCol
    lastRow = WS.Cells(WS.Rows.Count, i).End(xlUp).Row
    If FindLastRow < lastRow Then
        FindLastRow = lastRow
    End If
Next i
If FindLastRow < 17 Then FindLastRow = 17
End Function

Function AddDesignFormatting(WS As Worksheet, fontColorTarget As Range, interiorColorTarget As Range, Optional ByVal fontColor As Long, Optional ByVal interiorColor As Long)

If WS.Name = "M3 Upload Template" Then
    fontColorTarget.Font.Color = fontColor
    interiorColorTarget.Interior.Color = interiorColor
End If

End Function
Function RegKeyExists(key)
  Dim oShell, entry
  On Error Resume Next
 
  Set oShell = CreateObject("WScript.Shell")
  entry = oShell.RegRead(key)
  If Err.Number = 0 Then
    Err.Clear
    RegKeyExists = True
  Else
    Err.Clear
    RegKeyExists = False
  End If
  On Error GoTo 0
End Function
Function CheckNoOfRecordsStillToProcess(ByVal StartRow As Long, ByVal EndRow As Long, ByRef xSheet As Worksheet) As Long
'Check to see if there any records still to process if NOK or Blank cells
Dim NOK As Long
Dim Blank As Long

With xSheet.Application.WorksheetFunction
    NOK = .CountIf(Range("A" & StartRow & ":A" & EndRow), "NOK")
    Blank = .CountIf(Range("A" & StartRow & ":A" & EndRow), "")
End With

CheckNoOfRecordsStillToProcess = NOK + Blank
End Function
Function sheetExists(sheetToFind As String) As Boolean

Dim Sheet As Worksheet
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Sub RemoveValidation()
  Dim WS As Worksheet
  For Each WS In ActiveWorkbook.Worksheets
    If WS.Name = WS.Range("B3").value & " - " & WS.Range("B5").value Then
        WS.Cells.Validation.Delete
    End If
  Next WS
End Sub

Function CreateResponseSheet(sheetName As String) As String
    Dim oldSheet As Worksheet
    Set oldSheet = ActiveSheet
    Application.ScreenUpdating = False
    If sheetExists(Left(sheetName, 31)) = False Then
        ActiveWorkbook.Sheets.Add(after:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = Left(sheetName, 31)
        CreateResponseSheet = Left(sheetName, 31)
    Else
        'ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)).Name = Left(sheetName, 29) & ActiveWorkbook.Sheets.Count + 1
        CreateResponseSheet = Left(sheetName, 31)
    End If
    oldSheet.Activate
    Application.ScreenUpdating = True
End Function
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
End Function
Sub UploadSummary()
Dim WS As Worksheet
Dim strMsg As String
Set WS = ActiveSheet

With WS
    strMsg = UploadStatus(WS, .Range("B7").value, .Range("B8").value, "Upload Status:")
End With
If strMsg <> vbNullString Then
    MsgBox strMsg, vbInformation, "Summary"
Else
    MsgBox "No data in Start Row and End Row in the M3 Template", vbInformation, "Upload Summary"
End If
End Sub

Public Sub FindLastRowCol(ByRef Wsht As Worksheet, ByRef RowLast As Long, _
                          ByRef ColLast As Long)

  ' Sets RowLast and ColLast to the last row and column with a value
  ' in worksheet Wsht

  ' The motivation for coding this routine was the discovery that Find by
  ' previous row found a cell formatted as Merge and Center but Find by
  ' previous column did not.
  ' I had known the Find would miss merged cells but this was new to me.

  ' Dec16  Coded
  ' Corrected handling of UserRange
  ' SpecialCells was giving a higher row number than Find for
  '    no reason I could determine.  Added code to check for a
  '    value on rows and columns above those returned by Find
  ' Found column with value about that found by Find

  Dim ColCrnt As Long
  Dim ColLastFind As Long
  Dim ColLastOther As Long
  Dim ColLastTemp As Long
  Dim ColLeft As Long
  Dim ColRight As Long
  Dim Rng As Range
  Dim RowIncludesMerged As Boolean
  Dim RowBot As Long
  Dim RowCrnt As Long
  Dim RowLastFind As Long
  Dim RowLastOther As Long
  Dim RowLastTemp As Long
  Dim RowTop As Long

  With Wsht

    Set Rng = .Cells.Find("*", .Range("A1"), xlFormulas, , xlByRows, xlPrevious)
    If Rng Is Nothing Then
      RowLastFind = 0
      ColLastFind = 0
    Else
      RowLastFind = Rng.Row
      ColLastFind = Rng.Column
    End If

    Set Rng = .Cells.Find("*", .Range("A1"), xlValues, , xlByColumns, xlPrevious)
    If Rng Is Nothing Then
    Else
      If RowLastFind < Rng.Row Then
        RowLastFind = Rng.Row
      End If
      If ColLastFind < Rng.Column Then
        ColLastFind = Rng.Column
      End If
    End If

    Set Rng = .Range("A1").SpecialCells(xlCellTypeLastCell)
    If Rng Is Nothing Then
      RowLastOther = 0
      ColLastOther = 0
    Else
      RowLastOther = Rng.Row
      ColLastOther = Rng.Column
    End If

    Set Rng = .UsedRange
    If Rng Is Nothing Then
    Else
      If RowLastOther < Rng.Row + Rng.Rows.Count - 1 Then
        RowLastOther = Rng.Row + Rng.Rows.Count - 1
      End If
      If ColLastOther < Rng.Column + Rng.Columns.Count - 1 Then
        ColLastOther = Rng.Column + Rng.Columns.Count - 1
      End If
    End If

    If RowLastFind < RowLastOther Then
      ' Higher row found by SpecialCells or UserRange
      Do While RowLastOther > RowLastFind
        ColLastTemp = .Cells(RowLastOther, .Columns.Count).End(xlToLeft).Column
        If ColLastTemp > 1 Or .Cells(RowLastOther, 1).value <> "" Then
          Debug.Assert False
          ' Is this possible?
          ' Row after RowLastFind has value
          RowLastFind = RowLastOther
          Exit Do
        End If
        RowLastOther = RowLastOther - 1
      Loop
    ElseIf RowLastFind > RowLastOther Then
      Debug.Assert False
      ' Is this possible?
    End If
    RowLast = RowLastFind

    If ColLastFind < ColLastOther Then
      ' Higher column found by SpecialCells or UserRange
      Do While ColLastOther > ColLastFind
        RowLastTemp = .Cells(.Rows.Count, ColLastOther).End(xlUp).Row
        If RowLastTemp > 1 Or .Cells(1, ColLastOther).value <> "" Then
          'Debug.Assert False
          ' Column after ColLastFind has value
          ' Possible causes:
          '   * Find does not recognise merged cells
          '   * Find does not examine hidden cells
          ColLastFind = ColLastOther
          Exit Do
        End If
        ColLastOther = ColLastOther - 1
      Loop
    ElseIf ColLastFind > ColLastOther Then
      Debug.Assert False
      ' Is this possible
    End If
    ColLast = ColLastFind

  End With

End Sub


