Attribute VB_Name = "editTable"
Option Explicit
Public WS As Worksheet

Sub editTable()
'
Please note: you need to import ConfigProgressCode module and ProgressDialogue form in order to work
'
Dim questionBoxPopUp As VbMsgBoxResult
Dim currentProgressBar As New ProgressDialogue
Dim i As Long
i = 0
    questionBoxPopUp = MsgBox("Are you sure you want to edit this worksheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit table report?")
    If questionBoxPopUp = vbNo Then Exit Sub

    currentProgressBar.Configure "Editing... " & "Please wait!", "Editing...", i, 100, , True, True

On Error GoTo ErrorHandler

    currentProgressBar.Show
    currentProgressBar.SetValue i
    currentProgressBar.SetStatus "Editing..."
    If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

Application.ScreenUpdating = False

    Range(Cells(1, 1), Cells(Rows.Count, 1).End(xlUp)).Offset(1, 0).Select
    ActiveWindow.FreezePanes = True
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :=";", TrailingMinusNumbers:=True


        Call removeHeaderSpace
        i = i + 25
        currentProgressBar.SetValue i
        currentProgressBar.SetStatus "Editing... Removing the header space"

        Call removeProductCodeSpace
         i = i + 25
        currentProgressBar.SetValue i
        currentProgressBar.SetStatus "Editing... Removing the product space"

        changeProperties WS
         i = i + 25
        currentProgressBar.SetValue i
        currentProgressBar.SetStatus "Editing... Formatting table and changing printing properties"
        Call AutoSumColumnAF


    Range(Cells(1, 23), Cells(Rows.Count, 23).End(xlUp)).Offset(1, 0).Select

        i = i + 25
        currentProgressBar.SetValue i
        currentProgressBar.SetStatus "Saving... Getting folder path and saving file"

       'Call SaveAsToFolderPath 'This line of code is disabled until the path of the folder is changed



         Application.ScreenUpdating = True

          Unload currentProgressBar

       MsgBox "Process completed!", vbInformation

    Exit Sub
CanceledBtnPressed:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Cancelled By User.", vbInformation

    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print Err.Description
    MsgBox "Sorry, an error occured!" & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"

End Sub

Private Function SelectFirstEmptyRowInColumn(ByVal WS As Worksheet, Optional ByVal fromColumn As Long = 1) As Long

    With WS
            SelectFirstEmptyRowInColumn = .Cells(.Rows.Count, 32).End(xlUp).Row + 1
        End With
End Function
 Private Sub AutoSumColumnAF()

Dim selectLastRow As Long

    selectLastRow = SelectFirstEmptyRowInColumn(ActiveSheet, 32)
        Cells(selectLastRow, 32).Select
        With ActiveCell
        .Formula = WorksheetFunction.Sum(Range(Cells(2, .Column), ActiveCell))
        .Font.Bold = True
        .Font.ColorIndex = 3
        End With


 End Sub
Private Sub changeProperties(WS As Worksheet)

Dim tableRng As Range
Dim columnsToHide As Range
Dim headerFormat As Range

Set headerFormat = Range(Cells(1, 1), Cells(1, Columns.Count).End(xlToLeft)).Columns
Set tableRng = Range("A2").CurrentRegion
Set WS = ActiveSheet

With tableRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With headerFormat.Font
                      .Name = "Arial"
                      .FontStyle = "Bold"


    End With
    'headerFormat.Interior.ColorIndex = 17 'Another way to change the header colour
    headerFormat.Interior.Color = RGB(121, 171, 251) 'Current colour for header blue
    headerFormat.AutoFilter
    tableRng.Cells.EntireColumn.AutoFit
With WS

   Set columnsToHide = Application.Union(.Columns("A:F"), _
                                         .Columns("H:O"), _
                                         .Columns("Q"), _
                                         .Columns("S:V"), _
                                         .Columns("Z:AE"), _
                                         .Columns("AI:BD"))
        columnsToHide.EntireColumn.Hidden = True

   End With


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

    Private Sub removeHeaderSpace()
'
' Working version with no errors.

    Dim rngRemoveSpace As Range
    Dim CellChecker As Range


  Set rngRemoveSpace = Range(Cells(1, 1), Cells(2, Columns.Count).End(xlToLeft)).Columns

    rngRemoveSpace.Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False

    For Each CellChecker In rngRemoveSpace.Cells
        CellChecker.value = Application.Trim(CellChecker.value)
        CellChecker.value = Application.Clean(CellChecker.value)
    Next CellChecker

    Set rngRemoveSpace = Nothing

End Sub

    Private Sub removeProductCodeSpace()
'
' Working version with no errors.

    Dim rngRemoveSpace As Range
    Dim CellChecker As Range


  Set rngRemoveSpace = Range(Cells(1, 23), Cells(Rows.Count, 23).End(xlUp)).Offset(1, 0)

    rngRemoveSpace.Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False

    For Each CellChecker In rngRemoveSpace.Cells
        CellChecker.value = Application.Trim(CellChecker.value)
        CellChecker.value = Application.Clean(CellChecker.value)
    Next CellChecker

    Set rngRemoveSpace = Nothing

End Sub

 Private Sub SaveAsToFolderPath()
'
'
Dim MyFileName As String
Dim folderPath As String
Dim dateFormat As String
Dim saveDetails As String

'Change file path to where you want to save the file please note that the folder needs to exist before this line of code is executed
folderPath = "C:\Users\userName\Desktop\userFolder\Table report\"

       dateFormat = Format(Now, "dd.mm.yyyy HH-mm-ss AMPM")
       MyFileName = Range("G2").value
       saveDetails = folderPath & MyFileName & " - Next Delivery " & dateFormat & ".xlsx"

    If Not ActiveWorkbook.Saved Then
        ActiveWorkbook.SaveAs saveDetails, xlWorkbookDefault
    End If

End Sub
