Attribute VB_Name = "EditReportOld"
Option Explicit
Public Sub editReportOld()
'
'
'
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim WS As Worksheet

    questionBoxPopUp = MsgBox("Are you sure you want to edit Daily Report?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit Daily Report?")
    If questionBoxPopUp = vbNo Then Exit Sub

       On Error GoTo ErrorHandler

            For Each WS In ActiveWorkbook.Worksheets
            If WS.Name <> "SheetName6" Then '<--Ignore this sheet
                WS.Activate
                 Application.ScreenUpdating = False
                    editingProperties WS
                 Application.ScreenUpdating = True
        End If

                Next WS

            Application.ScreenUpdating = True

            MsgBox "Process completed!", vbInformation
            'MsgBox "Please note:" & vbNewLine & vbNewLine & "1. You will be redirected to print preview page." & vbNewLine & "2. Proceed with printing all reports.", vbInformation

           'Worksheets.PrintPreview 'To activate this line when happy to print everything

            Sheets(1).Select 'Please note that if this line is not executed you need to manually deselect the sheets by clicking on a different one!!!


        Exit Sub '<--- exit here if no error occured
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print Err.Description
        MsgBox "Sorry, an error occured." & vbCrLf & Err.Description, vbCritical, "Error!"

    End Sub

    Private Sub editingProperties(WS As Worksheet)

With WS
       .Range("A1:B5").Copy
       .Range("B1:C5").PasteSpecial
       .Range("B1:C2").Select
        Selection.Merge
       .Range("B1:C2").Font.Size = 24
       .Range("B4").Font.Size = 16
        ActiveWindow.Split = False
       .Cells.EntireColumn.AutoFit
       .Cells.EntireRow.AutoFit
       .Columns("A").EntireColumn.Hidden = True

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
            .FitToPagesTall = False 'switched off in case report bigger than 1 page it will scale down to unreadable format

        End With
    End Sub
