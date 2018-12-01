Attribute VB_Name = "SpotCheck"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub EditSpotCheck()

    Dim WS As Worksheet
    Dim delRange As Range
    Dim lRow As Long, i As Long
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim currentProgressBar As New ProgressDialogue

    questionBoxPopUp = MsgBox("Are you sure you want to edit the spot check worksheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit spot check report")
    If questionBoxPopUp = vbNo Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Call CopySheet
    
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
        
        Set delRange = Nothing

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
        
        Set delRange = Nothing

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
        .Columns("C").ColumnWidth = 9
        .Columns("K").ColumnWidth = 40
        
    End With

    Range("A5").EntireRow.AutoFit
    EditPrintProperties WS
    Call CreateTable
    Call ExtractSizes
    Call SaveAsToFolderPath
    
ScreenUpdate:
    Application.ScreenUpdating = True
    'ws.Range("A5").Activate
    'Selection.AutoFilter
    WS.Range("A6").Activate
    ActiveWindow.FreezePanes = True
    
    Unload currentProgressBar
    MsgBox "Process completed!", vbInformation, Title:="Spot check report"
    
    Exit Sub
CanceledBtnPressed:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Cancelled By User.", vbInformation
  
    Exit Sub '<--- exit here if no error occured
ErrorHandler:
    'Debug.Print Err.Number; Err.Description
    MsgBox "Sorry, an error occured." & vbNewLine & vbNewLine & "Please print screen with the error message together with step by step commands that triggered the error to the developer in order to fix it." & vbNewLine & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"
 Resume ScreenUpdate
End Sub
Private Sub CopySheet()
 
    Dim MySheetName As String
    MySheetName = "Edited_Spot_Check"
    Dim i As Integer

        If sheetExists("Edited_Spot_Check") Then
            For i = 1 To Worksheets.Count
                If Worksheets(i).Name Like "*Variance*" Then
            
                Sheets(i).Copy before:=Sheets(i)
                ActiveSheet.Name = MySheetName & Worksheets.Count
                Sheets(i).Tab.Color = RGB(31, 237, 139)
                End If
            Next i
        Else

            Sheets(1).Copy before:=Sheets(1)
            ActiveSheet.Name = MySheetName
        
            Sheets(1).Tab.Color = RGB(255, 10, 10)
            Sheets(2).Tab.Color = RGB(31, 237, 139)

    End If

End Sub
Private Function sheetExists(sheetToFind As String) As Boolean

Dim Sheet As Worksheet
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
Private Sub EditPrintProperties(WS As Worksheet)

Dim lastRow As Long

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
 
    lastRow = Range("G" & Rows.Count).End(xlUp).Row
    ActiveSheet.PageSetup.PrintArea = "A1:K" & lastRow
    
End Sub
Private Sub ExtractSizes()

Dim strMain As String
Dim str1 As String
Dim str2 As String
Dim sizeToExtract As String
Dim i, x, y As Long
Dim sizeCode As Long
Dim temp As Variant

str1 = "/"
str2 = "/"
 
        i = 6 'start on row i
        Do While Not IsEmpty(Cells(i, 4)) 'do until cell is empty
            strMain = Cells(i, 6)
            x = InStr(1, strMain, str1)
            y = InStr(1, strMain, str2)
        
        If Abs(y - x) < Len(str1) Then
            y = InStr(x + Len(str1), strMain, str2)
                If x = y Then 'try to search 2nd half of string for unique match
                    y = InStr(x + 1, strMain, str2)
                End If
        End If

        If x = 0 And y = 0 Then GoTo ErrorHandler:
            If y = 0 Then
                y = Len(strMain) + Len(str2) 'just to make it arbitrarily large
                If x = 0 Then
                    x = Len(strMain) + Len(str1) 'just to make it arbitrarily large
                 End If
             End If
     
        If x > y And y <> 0 Then 'swap order
            temp = y
            y = x
            x = temp
            temp = str2
            str2 = str1
            str1 = temp
        End If

        sizeCode = Cells(i, 3)
        Select Case sizeCode
        Case 99 '--> if 99 in cell then
            Cells(i, 3) = "NOSIZ"
        Case 4601
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "0/3"
        Case 4602
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "3/6"
        Case 4603
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "6/9"
        Case 4605
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "9/12"
        Case 4606
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "12/18"
        Case 4607
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "18/24"
        Case 4801
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "15/16"
        Case 4802
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "17/18"
        Case 4805
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "19/21"
        Case 4806
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "22/24"
        Case 4807
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "25/27"
        Case 4809
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "25/27"
        Case 4810
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "28/30"
        Case 4811
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "40/42"
        Case 4817
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "31/33"
        Case 4818
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "34/37"
        Case 4819
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "38/40"
            
        Case 9223
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "26/27"
        Case 9224
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "30/31"
        Case 9228
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "15/16"
        Case 9229
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "17/18"
        Case 9230
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "19/20"
        Case 9231
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "21/22"
        Case 9232
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "19/20"
        Case 9233
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "21/22"
        Case 9234
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "15/18"
        Case 9235
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "19/22"
        Case 9236
            Cells(i, 3).NumberFormat = "@"
            Cells(i, 3) = "23/26"
            
        Case Else
                x = x + Len(str1)
                sizeToExtract = Trim(Mid(strMain, x, y - x))
                Cells(i, 3) = sizeToExtract
        End Select
        
        i = i + 1 'increment row
    Loop
    
Exit Sub

ErrorHandler:
MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & Err.Number & " " & Err.Description & vbNewLine & vbCrLf & "Aborting", , "Strings not found"

End Sub
Private Sub CreateTable()
    Dim lo As ListObject
    
    If Not TableExistsOnSheet(ActiveSheet, ActiveSheet.Name) Then
        Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, [A5].CurrentRegion, , xlYes)
        With lo
            .Name = ActiveSheet.Name
            .TableStyle = "TableStyleMedium23"
        End With
    End If
 Set lo = Nothing
End Sub
Private Function TableExistsOnSheet(WS As Worksheet, sTableName As String) As Boolean
'--> Note this method will fail if the name of the sheet contains the name with space or ()
    TableExistsOnSheet = WS.Evaluate("ISREF(" & sTableName & ")")
End Function
Private Sub SaveAsToFolderPath()
'Set reference to Microsoft Scripting RunTime to see the properties and methods available in the IntelliSense
'The below 2 example will display the IntelliSense if the reference is set on
'Dim fso As Scripting.FileSystemObject
'Set fso = New Scripting.FileSystemObject

    Dim myFileName As String
    Dim newFolderPath As String
    Dim dateFormat As String
    Dim saveDetails As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Change file path to where you want to save the file
     newFolderPath = Environ("UserProfile") & "\Desktop\Marius\STOCK TAKE\Brompton\MINI STOCK TAKE\2018\"
    
        If Not fso.FolderExists(newFolderPath) Then
               fso.CreateFolder newFolderPath
        End If
        
              dateFormat = Format(Now, " dd.mm.yyyy")
              myFileName = Mid(Range("A3").value, 13, 50)
              saveDetails = newFolderPath & dateFormat & " - " & myFileName & ".xlsx"
              
        If Not fso.FileExists(saveDetails) Then
              
           If Not ActiveWorkbook.Saved Then
               ActiveWorkbook.SaveAs saveDetails, xlWorkbookDefault
           End If
        End If
        
    Set fso = Nothing
End Sub
Sub ComparingFrozenReportWithConsReport()
'--> Need to add dependency SmartUtlilities

    'Source list
    Dim allOnHandList As Range
    Dim allOnHandListResult As Range
    Dim allOnHandListSkuCell As Range
    Dim allOnHandListHeaderRowsCount As Integer
    Dim allOnHandListHeaderRowNum As Integer
    Dim foundMatchingallOnHandListSku As Range
    
    'Comparing against list
    Dim consList As Range
    Dim consListCellRow As Long
    Dim LBoundRow As Long
    Dim UBoundRow As Long
    Dim consListHeaderRowNum As Integer
    
    'Variable to hold if match found on the comparing against list
    Dim allOnHandListSkuCriteria As Variant
    
     On Error GoTo ErrorHandler
    
    '---> this method needs all data to be visible in order to loop through all cells
    SmartUtilities.ResetFilters
    
    '---> Allows users to select the ranges in case the table columns will change in the future
    Set allOnHandList = Application.InputBox("Select your all on hand list list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not allOnHandList Is Nothing Then
            If allOnHandList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the all on hand list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
     Set allOnHandListResult = Application.InputBox("Select the column header cell in the all on hand list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not allOnHandListResult Is Nothing Then
            If allOnHandListResult.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple cells selected! Please pick only the header cell in the all on hand list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
        
      Set consList = Application.InputBox("Select your consignment list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not consList Is Nothing Then
            If consList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the consignment sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
    allOnHandListHeaderRowsCount = allOnHandListResult.Row
    allOnHandListHeaderRowNum = allOnHandListResult.Row - 1
    
    consListHeaderRowNum = consList.Row - 1
      
        Application.ScreenUpdating = False
        
    '---> Allows users to compare the consignment list to the inventory list in order to find matches
    For Each allOnHandListSkuCell In allOnHandList
        allOnHandListSkuCriteria = Trim(allOnHandListSkuCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
        With consList
            Set foundMatchingallOnHandListSku = .Find(What:=allOnHandListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, _
                                                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, _
                                                SearchFormat:=False) 'finds a match
        End With
           
    If foundMatchingallOnHandListSku Is Nothing Then
        If allOnHandListSkuCell.Row = allOnHandListHeaderRowsCount Then
            If allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value = vbNullString Then
                allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value = "Open Consignment List"
                allOnHandListResult.Font.FontStyle = "Bold"
            End If
        End If
        
      Else
    
    
consListCellRow = 1
NextIteration:
'Check multiple rows for the same value in the range
    With consList
            If consListCellRow = 1 Then
                    consListCellRow = .Find(What:=allOnHandListSkuCriteria, After:=.Cells(consListCellRow, 1), LookIn:=xlValues, _
                                                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, _
                                                SearchFormat:=False).Row
                    LBoundRow = consListCellRow
                
                Else
                    consListCellRow = .Find(What:=allOnHandListSkuCriteria, After:=.Cells(consListCellRow, 1), LookIn:=xlValues, _
                                                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, _
                                                SearchFormat:=False).Row
                    UBoundRow = consListCellRow
            End If
    End With
        
    If LBoundRow = UBoundRow Then GoTo NextSkuCell:
        
    If allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value = vbNullString Then
    
         allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value = consList.Cells(consListCellRow - consListHeaderRowNum, 6).value & " on cons " _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, -7).value _
                                                                        & " (" _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, -6).value _
                                                                        & ") " _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, -5).value
        Else
        
         allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value = allOnHandListResult.Cells(allOnHandListSkuCell.Row - allOnHandListHeaderRowNum).value _
                                                                        & " // " _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, 6).value & " on cons " _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, -7).value _
                                                                        & " (" _
                                                                        & consList.Cells(consListCellRow - consListHeaderRowNum, -6).value _
                                                                        & ") " _
                                                                       & consList.Cells(consListCellRow - consListHeaderRowNum, -5).value
        
        End If
    End If
    
    If LBoundRow <> UBoundRow Then GoTo NextIteration:
         
NextSkuCell:
    LBoundRow = 0
    UBoundRow = 0
    
    Next allOnHandListSkuCell
     
        Application.ScreenUpdating = True
        
        MsgBox "Process completed!"
    
ErrorHandler:
        Application.ScreenUpdating = True
        Select Case Err.Number
                Case 424
                Exit Sub
                Case 0
                Exit Sub
                
                Case Else
                Debug.Print Err.Number, Err.Description
                MsgBox Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
                
        End Select

End Sub



