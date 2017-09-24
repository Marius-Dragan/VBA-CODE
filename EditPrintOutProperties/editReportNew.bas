Attribute VB_Name = "editReportNew"
Option Explicit
Public Sub editReportNew()
'
Please note: you need to import ConfigProgressCode module and ProgressDialogue form in order to work
'
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim WS As Worksheet
    Dim WS_Count As Long
    Dim i As Long
    Dim currentProgressBar As New ProgressDialogue

    Const staffList As String = "SheetName1, SheetName2, SheetName3, SheetName4, SheetName5"

    Application.ScreenUpdating = False

    WS_Count = ActiveWorkbook.Worksheets.Count - 1
    currentProgressBar.Configure "Editing... " & "Please wait!", "Editing...", i, WS_Count, , True, True



    questionBoxPopUp = MsgBox("Are you sure you want to edit Daily Report?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit Daily Report?")
    If questionBoxPopUp = vbNo Then Exit Sub

       On Error GoTo ErrorHandler
        currentProgressBar.Show
            For i = 1 To WS_Count
                currentProgressBar.SetValue i
                currentProgressBar.SetStatus "Editing and printing " & Worksheets(i).Name & " worksheet " & i & " out of " & WS_Count & " done "
                If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

                    Worksheets(i).Activate

                        If Worksheets(i).Name <> "SheetName6" Then '<--Ignore sheet

                            editingProperties WS

                            If IsError(Application.Match(WS.Name, Split(staffList, ", "), 0)) Then
                                Debug.Print WS.Name & " is not included in the staff list!"

                            Else:
                                Debug.Print WS.Name & " is included in the staff list!"
                                WS.PrintOut ' prints out every page that is editing
                            End If

                        End If


            Next i

               Unload currentProgressBar

            Application.ScreenUpdating = True

            MsgBox "Process completed!", vbInformation

Exit Sub
CanceledBtnPressed:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Cancelled By User.", vbInformation

Exit Sub '<--- exit here if no error occured
ErrorHandler:
 Application.ScreenUpdating = True
    Debug.Print Err.Description
        MsgBox "Sorry, an error occured." & vbCrLf & Err.Description, vbCritical, "Error!"

    End Sub

    Private Sub editingProperties(WS As Worksheet)

Set WS = Application.ActiveSheet
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
            .FitToPagesTall = False

        End With
    End Sub
