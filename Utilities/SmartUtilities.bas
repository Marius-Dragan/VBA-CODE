Attribute VB_Name = "SmartUtilities"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub ResetFilters()

   Dim WS As Worksheet
   Dim WB As Workbook
   Dim listObj As ListObject
      
    Set WB = ActiveWorkbook
'This is if you place the macro in your personal wb to be able to reset the filters on any wb you're currently working on. Remove the set wb = thisworkbook if that's what you need
           For Each WS In WB.Worksheets
              If WS.AutoFilterMode Then
                 WS.AutoFilter.ShowAllData 'clears filters from the sheet
              Else
'This removes "normal" filters in the workbook - however, it doesn't remove table filters
              End If
                For Each listObj In WS.ListObjects 'And this removes table filters. You need both aspects to make it work.
                     If listObj.ShowAutoFilter Then
                          listObj.AutoFilter.ShowAllData 'clears filters from the table
                          'listObj.Range.AutoFilter 'To set or unset the filter for the table
                          listObj.Sort.SortFields.Clear
                     End If
                Next listObj
            Next WS
End Sub
Sub OptimizeVBA(isOn As Boolean)

    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub
Sub TrimCellsArrayMethodUsedRange()

Dim arrData() As Variant
Dim arrReturnData() As Variant
Dim rng As Excel.Range
Dim lRows As Long
Dim lCols As Long
Dim i As Long, j As Long
Dim WS As Worksheet

  Set WS = ActiveSheet
  lRows = WS.UsedRange.Rows.Count
  lCols = WS.UsedRange.Columns.Count

  ReDim arrData(1 To lRows, 1 To lCols)
  ReDim arrReturnData(1 To lRows, 1 To lCols)
    
    
  Set rng = WS.UsedRange
  arrData = rng.value

  For j = 1 To lCols
    For i = 1 To lRows
      arrReturnData(i, j) = Trim(arrData(i, j))
    Next i
  Next j

  rng.value = arrReturnData

  Set rng = Nothing
End Sub
Sub TrimCellsArrayMethodSelection()

Dim arrData() As Variant
Dim arrReturnData() As Variant
Dim rng As Excel.Range
Dim lRows As Long
Dim lCols As Long
Dim i As Long, j As Long

  lRows = Selection.Rows.Count
  lCols = Selection.Columns.Count

  ReDim arrData(1 To lRows, 1 To lCols)
  ReDim arrReturnData(1 To lRows, 1 To lCols)

  Set rng = Selection
  arrData = rng.value

  For j = 1 To lCols
    For i = 1 To lRows
      arrReturnData(i, j) = Trim(arrData(i, j))
    Next i
  Next j

  rng.value = arrReturnData

  Set rng = Nothing
End Sub

Sub RemoveSpaceV2()
'
' Working version with no errors.

    Dim rngRemoveSpace As Range
    Dim CellChecker As Range
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
     'On Error GoTo ErrorHandler
     
  Set rngRemoveSpace = Intersect(ActiveSheet.UsedRange, Selection)
    rngRemoveSpace.Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
    
    For Each CellChecker In rngRemoveSpace.Cells
        CellChecker.value = Application.Trim(CellChecker.value)
        CellChecker.value = Application.Clean(CellChecker.value)
        
    Next CellChecker
    
Application.ScreenUpdating = True

    Set rngRemoveSpace = Nothing
'ErrorHandler:
   'Debug.Print "Error number: " & Err.Number & " " & Err.Description
'        MsgBox "Sorry, an error occured." & vbCrLf & Err.Description, vbCritical, "Error!"
    
End Sub
Private Sub LoopThroughFiles()

    Dim folderPath As String
    Dim filename As String
    Dim WB As Workbook
  
    folderPath = "C:\Users\str-brompton.smc.uk\Desktop\New folder\" 'change to suit
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath + "\"
    
    filename = Dir(folderPath & "*.xls")
    Do While filename <> ""
    Debug.Print (filename)
      Application.ScreenUpdating = False
        Set WB = Workbooks.Open(folderPath & filename)
         
        'Call a subroutine here to operate on the just-opened workbook
        'Call Edit121ReportWithAutoPrint
        WB.Close False
        filename = Dir
    Loop
  Application.ScreenUpdating = True
End Sub
Sub SortActiveWorksheets()


Dim i As Integer
Dim j As Integer
Dim questionBoxPopUp As VbMsgBoxResult


'
' Prompt the user as which direction they wish to
' sort the worksheets.
'
   questionBoxPopUp = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
     & "Clicking No will sort in Descending Order.", _
     vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
   For i = 1 To Sheets.Count
      For j = 1 To Sheets.Count - 1
'
' If the answer is Yes, then sort in ascending order.
'
         If questionBoxPopUp = vbYes Then
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
'
' If the answer is No, then sort in descending order.
'
         ElseIf questionBoxPopUp = vbNo Then
            If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
         End If
      Next j
   Next i
    MsgBox "Process completed!", vbInformation
End Sub
Private Sub CopySheetsToNewWorkbook()

Dim xPath As String
Dim xWs As Worksheet
Dim questionBoxPopUp As VbMsgBoxResult

 questionBoxPopUp = MsgBox("Are you sure you want to copy each worksheets as a new workbook in the current folder?", vbQuestion + vbYesNo + vbDefaultButton1, "Copy Worksheets?")
    If questionBoxPopUp = vbNo Then Exit Sub

On Error GoTo ErrorHandler
xPath = Application.ActiveWorkbook.Path

Application.ScreenUpdating = False
Application.DisplayAlerts = False

For Each xWs In ActiveWorkbook.Sheets
    xWs.Copy
    Application.ActiveWorkbook.SaveAs filename:=xPath & "\" & xWs.Name & ".xlsx"
    Application.ActiveWorkbook.Close False
Next xWs

Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox "Process completed!", vbInformation

 Exit Sub '<--- exit here if no error occured
ErrorHandler:
Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Debug.Print Err.Number; Err.Description
        MsgBox "Sorry, an error occured." & vbNewLine & vbNewLine & "Please print screen with the error message together with step by step commands that triggered the error to the developer in order to fix it." & vbNewLine & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"

End Sub
Sub AutoFit()

ActiveCell.CurrentRegion.EntireColumn.AutoFit
ActiveCell.CurrentRegion.EntireRow.AutoFit

End Sub

Private Sub EditPrintingProperties()

Set WS = Application.ActiveSheet
     With WS.PageSetup
            .PrintArea = ""
            .PrintTitleRows = ""
            .PrintTitleColumns = ""
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False

        End With
    End Sub

