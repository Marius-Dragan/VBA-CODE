Attribute VB_Name = "SpotCheck"
Option Explicit

Sub EditSpotCheck()
'
'Please note: you need to import ConfigProgressCode module and ProgressDialogue form in order to work
'

    Dim WS As Worksheet
    Dim delRange As Range
    Dim lRow As Long, i As Long
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim currentProgressBar As New ProgressDialogue

    questionBoxPopUp = MsgBox("Are you sure you want to edit the spot check worksheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit spot check file?")
    If questionBoxPopUp = vbNo Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    CopySheet

    Set WS = ActiveSheet

    With WS
        lRow = .Range("A" & .Rows.Count).End(xlUp).Row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lRow, , True, True
        currentProgressBar.Show

        '--> Delete All rows where Cell A and Cell B are empty
        For i = 6 To lRow

            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Looping and deleting all rows where cell A and cell B are empty " & i & " out of " & lRow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

            If Len(Trim(.Range("A" & i).value)) = 0 Or Len(Trim(.Range("B" & i).value)) = 0 Then
                If delRange Is Nothing Then
                    Set delRange = .Rows(i)
                Else
                    Set delRange = Union(delRange, .Rows(i))
                End If
            End If
        Next i

        If Not delRange Is Nothing Then delRange.Delete

        '--> Find the new last row
        lRow = .Range("A" & .Rows.Count).End(xlUp).Row

        '--> Insert a new column between G and H
        .Columns(8).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

        '--> Insert a formula =G6 & "(" & I6 & ")" in H6
        '--> Inserting the formula in the entire column in one go
        '--> and converting it to values
        .Range("H6:H" & lRow).Formula = "=G6 & ""("" & I6 & "")"""
        .Range("H6:H" & lRow).value = .Range("H6:H" & lRow).value
        '--> Copy the header from Col G to Col H so that we can delete the
        '--> Column G as it is not required anymore
        .Range("H5").value = .Range("G5").value
        .Columns(7).Delete
        .Range("K5").value = "Comments"

        currentProgressBar.Hide
        currentProgressBar.Show

        '--> Using a reverse loop to append values from bottom row to the row above
        '--> After appending clear the cell G so that we can later delete the row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lRow, , True, True
        For i = lRow To 7 Step -1

            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Using a reverse loop to append values from bottom row to the row above " & i & " out of " & lRow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

            If .Range("D" & i).value = .Range("D" & i - 1).value Then
                .Range("G" & i - 1).value = .Range("G" & i - 1).value & ", " & .Range("G" & i).value
                .Range("H" & i - 1).value = .Range("H" & i - 1).value + .Range("H" & i).value
                .Range("G" & i).ClearContents
            End If
        Next i

        Set delRange = Nothing

        currentProgressBar.Hide
        currentProgressBar.Show

        '--> Delete rows where Cell G is empty
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lRow, , True, True
        For i = 6 To lRow

            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Delete rows where the rows on cell G are empty " & i & " out of " & lRow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

            If Len(Trim(.Range("G" & i).value)) = 0 Then
                If delRange Is Nothing Then
                    Set delRange = .Rows(i)
                Else
                    Set delRange = Union(delRange, .Rows(i))
                End If
            End If
        Next i

        If Not delRange Is Nothing Then delRange.Delete

        '--> Find the new last row
        lRow = .Range("A" & .Rows.Count).End(xlUp).Row

        '--> Calculating the variance

        .Range("J6:J" & lRow).Formula = "=H6-I6"
        '.Range("J6:J" & lRow).Value = .Range("J6:J" & lRow).Value '<--- Line to convert formulas to values for column J


        With .Range("G" & lRow + 1)
        .value = "Grand Total:"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .BorderAround xlContinuous, xlThin
    End With

        With .Range("H" & lRow + 1)
        .Formula = "=SUM(H6" & ":H" & lRow & ")"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbRed
        .BorderAround xlContinuous, xlThin
    End With

        With .Range("I" & lRow + 1)
        .Formula = "=SUM(I6" & ":I" & lRow & ")"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbGreen
        .BorderAround xlContinuous, xlThin
    End With

        With .Range("J" & lRow + 1)
        .Formula = "=SUM(J6" & ":J" & lRow & ")"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbRed
        .BorderAround xlContinuous, xlThin
    End With

        .Range("G5:G" & "K5:K" & lRow).WrapText = False
        .Range("A5:K5").Interior.Color = RGB(141, 180, 227)
        .Cells(7).EntireColumn.AutoFit
        .Cells(11).EntireColumn.AutoFit
        .Columns("K").ColumnWidth = 30

    End With

    Range("A5").EntireRow.AutoFit
    editPrintProperties WS

ScreenUpdate:
    Application.ScreenUpdating = True
    WS.Range("A5").Activate
    Selection.AutoFilter
    WS.Range("A6").Activate
    ActiveWindow.FreezePanes = True

    Unload currentProgressBar

    MsgBox "Process completed!", vbInformation

    Exit Sub
CanceledBtnPressed:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Cancelled By User.", vbInformation

    Exit Sub '<--- exit here if no error occured
ErrorHandler:
    Debug.Print Err.Number; Err.Description
        MsgBox "Sorry, an error occured." & vbNewLine & vbNewLine & "Please print screen with the error message together with step by step commands that triggered the error to the developer in order to fix it." & vbNewLine & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"
        Resume ScreenUpdate
End Sub

  Private Sub CopySheet()

'
'
'

    Dim MySheetName As String
    MySheetName = "Edited Spot Check"

        If sheetExists("Edited Spot Check") = True Then
            MsgBox "Sheet named " & "'Edited Spot Check'" & " already exists. Please rename if you need another copy!", vbInformation, "Sheet exists!"
    End
        Else

             Sheets(1).Copy before:=Sheets(1)
            ActiveSheet.Name = MySheetName

        Sheets(1).Tab.Color = RGB(255, 10, 10)
     Sheets(2).Tab.Color = RGB(31, 237, 139)

    End If

End Sub

Function sheetExists(sheetToFind As String) As Boolean

'
'
'

Dim Sheet As Worksheet
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Private Sub editPrintProperties(WS As Worksheet)

'
'
'

Dim LastRow As Long

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

    LastRow = Range("G" & Rows.Count).End(xlUp).Row
    ActiveSheet.PageSetup.PrintArea = "A1:K" & LastRow

End Sub
