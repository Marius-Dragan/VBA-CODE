Attribute VB_Name = "editReportOldVersion"
Public Sub editReportOldVersion()
'
'
'
    Dim questionBoxPopUp As VbMsgBoxResult
Dim WS As Worksheet

    questionBoxPopUp = MsgBox("Are you sure you want to edit all sheets in this workbook?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit Workbook?")
    If questionBoxPopUp = vbNo Then Exit Sub

       On Error GoTo ErrorHandler

            For Each WS In ActiveWorkbook.Worksheets
            WS.Activate
            Application.ScreenUpdating = False
             editingProperties WS
             Application.ScreenUpdating = True
                Next WS

            Application.ScreenUpdating = True

            MsgBox "Please note:" & vbNewLine & vbNewLine & "1. You will be redirected to print preview page." & vbNewLine & "2. Proceed with printing all reports.", vbInformation

           Worksheets.PrintPreview
           'Worksheets.PrintOut 'This will automatically start printing.

            Sheets(1).Select 'Please note that if this line is not executed you need to manually deselect the sheets by clicking on a different one!!!


        Exit Sub '<--- exit here if no error occured
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print Err.Description
        MsgBox "Sorry, an error occured." & vbCrLf & Err.Description, vbCritical, "Error!"

    End Sub

    Private Sub editingProperties(WS As Worksheet)

Dim columnsToDelete As Range

With WS
       .Columns("A:F").UnMerge

    Set columnsToDelete = Application.Union(.Columns("B:C"), _
                                            .Columns("F:K"), _
                                            .Columns("P:R"), _
                                            .Columns("V:W"))
        columnsToDelete.Delete

       .Cells.EntireColumn.AutoFit
       .Range("A1:B2").Merge

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
