Attribute VB_Name = "Utilities"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub CountNonBlankCells()

Dim xTitleId As String
Dim rng As Range
Dim WorkRng As Range
Dim total As Long
On Error Resume Next
xTitleId = "SelectedCells"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each rng In WorkRng
    If Not IsEmpty(rng.value) Then
        total = total + 1
    End If
Next
MsgBox "There are " & total & " not blank cells in this range."
End Sub

Sub AutoSumToMsgBox()
     
    MsgBox "Sum of " & Selection.Address & " = " & Application.WorksheetFunction.Sum(Selection)
     
End Sub

Private Function SumOnlyVisible(r As Range) As Double
    Dim rCell As Range
     
    Application.Volatile
     
    Dim cell As Range
     
    For Each rCell In r.Cells
        With rCell
            If Not .Rows.Hidden And Not .Columns.Hidden Then SumOnlyVisible = SumOnlyVisible + .value
        End With
    Next
End Function

Private Function SumVisible(CRange As Object)
Dim TotalSum As Integer
Dim cell As Range
    Application.Volatile
    TotalSum = 0
    For Each cell In CRange
       If cell.Columns.Hidden = False Then
          If cell.Rows.Hidden = False Then
             TotalSum = TotalSum + cell.value
          End If
       End If
    Next
    SumVisible = TotalSum
End Function

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)

'To be put in ThisWorkbook for triger events to automatically change case
On Error Resume Next

If Intersect(Target, Range("E2:I500")) Is Nothing Then Exit Sub
Application.EnableEvents = False
Target.value = UCase(Target.value)
Application.EnableEvents = True

End Sub

Sub GenerateStyleFabricColour()

    Dim lastRow As Long
    Dim rowNum As Long
       
     Application.ScreenUpdating = False
    
    With ActiveSheet
        rowNum = Application.Max(1, .Cells(.Rows.Count, "M").End(xlUp).Row)
        lastRow = Application.Max(1, _
                    .Cells(.Rows.Count, "B").End(xlUp).Row, _
                    .Cells(.Rows.Count, "C").End(xlUp).Row, _
                    .Cells(.Rows.Count, "D").End(xlUp).Row)
           
            With .Cells(rowNum + 1, "B").Resize(lastRow - rowNum, 3).Select
                Call RemoveSpaceV2
                ActiveCell.Select
            End With
        With .Cells(lastRow - (lastRow - rowNum) + 1, "A").Resize(lastRow - rowNum, 1)
            .FormulaR1C1 = "=rc[1]&rc[2]&rc[3]"
            .value = .Value2
            .Offset(0, 1).value = .Value2
            .Offset(0, 2).ClearContents
            .Offset(0, 3).ClearContents
            .ClearContents
            .Offset(0, 5).Select
            ActiveCell.Select
        End With
    End With
    
    Application.ScreenUpdating = True
     
End Sub

Sub ChangeWeekdayColour()
Dim cell As Range
    
    For Each cell In Selection
            Select Case Weekday(cell.value)
            
                   Case Is = 1 'vbSunday (1)
                        cell.Columns("A").Interior.Color = RGB(89, 89, 89)
                        cell.Columns("C").Interior.Color = RGB(89, 89, 89)

                   Case Is = 2 'vbMonday (2)
                        cell.Columns("A").Interior.Color = RGB(166, 166, 166)
                        cell.Columns("C").Interior.Color = RGB(166, 166, 166)
                        
                   Case Is = 3 'vbTuesday (3)
                        cell.Columns("A").Interior.Color = RGB(166, 166, 166)
                        cell.Columns("C").Interior.Color = RGB(166, 166, 166)
    
                   Case Is = 4 'vbWednesday (4)
                        cell.Columns("A").Interior.Color = RGB(166, 166, 166)
                        cell.Columns("C").Interior.Color = RGB(166, 166, 166)
    
                   Case Is = 5 'vbThursday (5)
                        cell.Columns("A").Interior.Color = RGB(166, 166, 166)
                        cell.Columns("C").Interior.Color = RGB(166, 166, 166)
    
                   Case Is = 6 'vbFriday (6)
                        cell.Columns("A").Interior.Color = RGB(166, 166, 166)
                        cell.Columns("C").Interior.Color = RGB(166, 166, 166)
    
                   Case Is = 7 'vbSaturday (7)
                        cell.Columns("A").Interior.Color = RGB(89, 89, 89)
                        cell.Columns("C").Interior.Color = RGB(89, 89, 89)
                    
             End Select
        Next
End Sub
Sub RemoveAllSpaceInString()
    With ActiveSheet
        Call RemoveSpaceV2
        Intersect(Selection, .UsedRange).Replace " ", ""
    End With
End Sub


