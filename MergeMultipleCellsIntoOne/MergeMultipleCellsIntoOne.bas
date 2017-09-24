Attribute VB_Name = "MergeMultipleCellsIntoOne"
Option Explicit
Sub MergeMultipleCellsIntoOne()
'
'
    Dim lastRow As Long

     'OptimizeVBA True

    With Worksheets(2)
        lastRow = Application.Max(4, _
                    .Cells(.Rows.Count, "K").End(xlUp).Row, _
                    .Cells(.Rows.Count, "L").End(xlUp).Row, _
                    .Cells(.Rows.Count, "M").End(xlUp).Row)
            With .Cells(4, "K").Resize(lastRow - 4 + 1, 3).Select
                Call removeSpaceV2
        End With
           With .Cells(4, "H").Resize(lastRow - 4 + 1, 1)
            .FormulaR1C1 = "=rc[3]&rc[4]&rc[5]"
            .Value = .Value2
        End With
    End With

    With ActiveSheet
                     .Cells(1, "H").Resize(lastRow - 1 + 1, 7).Select
                     .PageSetup.PrintArea = Selection.Address
                     .Cells(4, "H").Resize(lastRow - 4 + 1, 1).Select
    End With

     'OptimizeVBA False

End Sub

Sub removeSpaceV2()
'

    Dim rngremovespace As Range
    Dim CellChecker As Range


  Set rngremovespace = Intersect(ActiveSheet.UsedRange, Selection)
    rngremovespace.Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False

    For Each CellChecker In rngremovespace.Cells
        CellChecker.Value = Application.Trim(CellChecker.Value)
        CellChecker.Value = Application.Clean(CellChecker.Value)
    Next CellChecker

    Set rngremovespace = Nothing

End Sub

Sub OptimizeVBA(isOn As Boolean)
'
'
    With Application
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not (isOn)
        .ScreenUpdating = Not (isOn)
        .DisplayAlerts = Not (isOn)
    End With
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub
