Attribute VB_Name = "SmartUtilities"
Option Explicit
Sub ResetFilters()
   Dim ws As Worksheet
   Dim wb As Workbook
   Dim listObj As ListObject
      
    Set wb = ActiveWorkbook
'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on. Remove the set wb = thisworkbook if that's what you need
           For Each ws In wb.Worksheets
              If ws.AutoFilterMode Then
                 ws.AutoFilter.ShowAllData 'clears filters from the sheet
              Else
'This removes "normal" filters in the workbook - however, it doesn't remove table filters
              End If
                For Each listObj In ws.ListObjects 'And this removes table filters. You need both aspects to make it work.
                     If listObj.ShowAutoFilter Then
                          listObj.AutoFilter.ShowAllData 'clears filters from the table
                          'listObj.Range.AutoFilter 'To set or unset the filter for the table
                          listObj.Sort.SortFields.Clear
                     End If
                Next listObj
            Next ws
End Sub
Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub
