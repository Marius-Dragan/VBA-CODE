Attribute VB_Name = "AMCFilter"
Sub advancedMultipleCriteriaFilter()
'
'
'
    Dim cellCriteria As Range, dataTable As Range, criteriaSelection As Range
    Dim filterCriteria() As String, filterFields() As Integer
    Dim i As Integer, criteriaToFilter As Range, rngSelection As Object

    On Error GoTo ErrorHandler

     Set rngSelection = Application.Selection
            If TypeOf rngSelection Is Range Then
             Set criteriaToFilter = rngSelection
             Else
            MsgBox "Cannot apply filter to your current selection as it is not a range! Please make another selection and try again." & vbNewLine & vbNewLine & "Note: selection can be a shape, chart, series and nothing!", vbInformation, "No filtering criteria selected!"
        Exit Sub
    End If
    
   
         Application.ScreenUpdating = False

    If criteriaToFilter.Rows.Count > 1 Then
            MsgBox "Cannot apply filter to multiple rows within the same column. Please make another selection and try again.", vbInformation, "Selection Error!"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    i = 1
    ReDim filterCriteria(1 To criteriaToFilter.Cells.Count) As String
    ReDim filterFields(1 To criteriaToFilter.Cells.Count) As Integer

    Set dataTable = criteriaToFilter.CurrentRegion
         For Each criteriaSelection In criteriaToFilter.Areas
             For Each cellCriteria In criteriaSelection
                 filterCriteria(i) = cellCriteria.Text
                 filterFields(i) = cellCriteria.Column - dataTable.Cells(1, 1).Column + 1
                  i = i + 1
        Next cellCriteria
    Next criteriaSelection

    With dataTable
        For i = 1 To UBound(filterCriteria)
            .AutoFilter field:=filterFields(i), Criteria1:=filterCriteria(i)
        Next i
    End With

    Call FirstEmptyRowSelection

    Set dataTable = Nothing
     Application.ScreenUpdating = True

    Exit Sub
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print "Error number: " & Err.Number & " " & Err.Description

End Sub

Sub resetFilters()
'
'

    On Error GoTo ErrorHandler

        Application.ScreenUpdating = False

            If ActiveSheet.FilterMode Then
         ActiveSheet.ShowAllData
   
        End If

    Application.ScreenUpdating = True

  Call FirstEmptyRowSelection

    Exit Sub
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print "Error number: " & Err.Number & " " & Err.Description


End Sub
Private Function SelectFirstEmptyRowInColumn(ByVal WS As Worksheet, Optional ByVal fromColumn As Long = 1) As Long

    With WS
            SelectFirstEmptyRowInColumn = .Cells(.Rows.Count, 8).End(xlUp).Row + 1
        End With
End Function
Private Sub FirstEmptyRowSelection()

Dim selectLastRow As Long

    selectLastRow = SelectFirstEmptyRowInColumn(ActiveSheet, 8)
        Cells(selectLastRow, 8).Select
  
 End Sub
 
 Sub resettingFiltersAndClearData()
'
'
'For the reset button on the worksheet

    On Error GoTo ErrorHandler

        Application.ScreenUpdating = False

            If ActiveSheet.FilterMode Then
         ActiveSheet.ShowAllData
   
        End If

     Range("A3:T3").ClearContents
    Application.ScreenUpdating = True

  Call FirstEmptyRowSelection

    Exit Sub
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print "Error number: " & Err.Number & " " & Err.Description


End Sub


