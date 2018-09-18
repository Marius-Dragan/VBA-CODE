Attribute VB_Name = "FlaggingReport"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub EditFlaggingReport()

    Dim WS As Worksheet
    Dim delRange As Range
    Dim lrow As Long, i As Long
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim currentProgressBar As New ProgressDialogue

    questionBoxPopUp = MsgBox("Are you sure you want to edit flagging list report?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit flagging list report")
    If questionBoxPopUp = vbNo Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Call CopySheet
    
    Set WS = ActiveSheet

    With WS
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        currentProgressBar.Show

        '--> Delete All rows where Cell A and Cell B are empty
        For i = 6 To lrow
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Looping and deleting all rows where cell A and cell B are empty " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:
            
            If Range("A" & i).value = "Style" Then
                Range("A" & i).ClearContents
            End If
            
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
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row
        
        
        .Columns("E").ColumnWidth = 18
        .Range("E5").value = "Style/Fabric/Colour"
        .Range("E6:E" & lrow).Formula = "=A6 & B6"
        .Range("E6:E" & lrow).value = .Range("E6:E" & lrow).value
        .Columns("C").ColumnWidth = 7
        Call ExtractSizes

        .Columns("G").ColumnWidth = 20
        .Range("G5").value = "All sizes on hand"
        
        .Columns(8).Delete
        .Columns(9).Delete
        
        .Columns("H").ColumnWidth = 15
        .Range("H5").value = "Total JDA Qty"
        .Columns("I").ColumnWidth = 30
        .Range("I5").value = "Comments"
        
        '--> Insert a formula =H6 & "(" & C6 & ")" in H6
        '--> Inserting the formula in the entire column in one go and converting it to values
        .Range("G6:G" & lrow).Formula = "=H6 & ""("" & C6 & "")"""
        .Range("G6:G" & lrow).value = .Range("G6:G" & lrow).value

        
        currentProgressBar.Hide
        currentProgressBar.Show
 

        '--> Using a reverse loop to append values from bottom row to the row above
        '--> After appending clear the cell G so that we can later delete the row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        For i = lrow To 7 Step -1
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Using a reverse loop to append values from bottom row to the row above " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:
            
            If .Range("E" & i).value = .Range("E" & i - 1).value Then
                .Range("G" & i - 1).value = .Range("G" & i - 1).value & ", " & .Range("G" & i).value
                .Range("H" & i - 1).value = .Range("H" & i - 1).value + .Range("H" & i).value
                .Range("G" & i).ClearContents
            End If
        Next i
        
        currentProgressBar.Hide
        currentProgressBar.Show

        '--> Delete rows where Cell G is empty
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        For i = 6 To lrow
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Delete rows where the rows on cell G are empty " & i & " out of " & lrow & " rows done"
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
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row


        With .Range("G" & lrow + 1)
        .value = "Grand Total:"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .BorderAround xlContinuous, xlThin
    End With
    
        With .Range("H" & lrow + 1)
        .Formula = "=SUM(H6" & ":H" & lrow & ")"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbRed
        .BorderAround xlContinuous, xlThin
    End With
        
        .Range("G5:G" & "C5:C" & lrow).WrapText = False
        .Range("A5:I5").Interior.Color = RGB(141, 180, 227)
        .Cells(3).EntireColumn.AutoFit
        .Cells(7).EntireColumn.AutoFit
        
    End With

    Range("A5").EntireRow.AutoFit
    EditPrintProperties WS
    Call CreateTable
    Call SaveAsToFolderPath
    
ScreenUpdate:
    Application.ScreenUpdating = True
    'ws.Range("A5").Activate
    'Selection.AutoFilter
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
 
    Dim MySheetName As String
    MySheetName = "All_Stock_On_Hand_Report"
    Dim i As Integer

        If sheetExists("All_Stock_On_Hand_Report") Then
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
Function sheetExists(sheetToFind As String) As Boolean

Dim Sheet As Worksheet
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
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
     newFolderPath = Environ("UserProfile") & "\Desktop\"
    
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
Private Sub EditSheet()
 
Dim i As Integer
Dim lrow As Long
Dim WS As Worksheet

        If sheetExists("All_Stock_On_Hand_Report") Then
            For i = 1 To Worksheets.Count
                If Worksheets(i).Name Like "*WOMENS*" Then
                    Worksheets(i).Activate
                    Call CreateTable
                    Worksheets(i).Tab.Color = RGB(31, 237, 139)
                ElseIf Worksheets(i).Name Like "*MENS*" Then
                    Worksheets(i).Activate
                    Call CreateTable
                    Worksheets(i).Tab.Color = RGB(31, 237, 139)
                ElseIf Worksheets(i).Name Like "*LICENSES*" Then
                    Worksheets(i).Activate
                    Call CreateTable
                    Worksheets(i).Tab.Color = RGB(31, 237, 139)
                ElseIf Worksheets(i).Name Like "*KIDS*" Then
                    Worksheets(i).Activate
                    Call CreateTable
                    Worksheets(i).Tab.Color = RGB(31, 237, 139)
                Else
                End If
            Next i
        Else
        
            For i = 1 To Worksheets.Count
                'If Worksheets(i).Name Like "*Sheet1*" Then
                If Worksheets(i).Name <> "WOMENS" And Worksheets(i).Name <> "MENS" And Worksheets(i).Name <> "LICENSES" And Worksheets(i).Name <> "KIDS" Then
                    Worksheets(i).Activate
                    Worksheets(i).Name = "All_Stock_On_Hand_Report"
                    Worksheets(i).Tab.Color = RGB(255, 10, 10)
                End If
            Next i
    End If

End Sub
Sub CompareFlaggingReportWithFrozenReport()
'--> Still to try exploring getting the range automatically from the tables name
'--> Need to add dependency SmartUtlilities

    'Source list
    Dim flaggingList As Range
    Dim flaggingListResult As Range
    Dim flaggingListStyleFabricColorCell As Range
    Dim flaggingListHeaderRowsCount As Integer
    Dim flaggingListHeaderRowNum As Integer
    
    'Comparing against list
    Dim currentAllSizesOnHandListCellRow As Long
    Dim AllSizesOnHandList As Range
    Dim foundMatchingAllSizesOnHandListStyleFabricColor As Range
    Dim AllSizesOnHandListHeaderRowNum As Integer

    'Variable to hold if match found on the comparing against list
    Dim flaggingListStyleFabricColorCriteria As Variant
    
    On Error GoTo ErrorHandler
    
    Call EditSheet
    
    Call SmartUtilities.ResetFilters
        
    '---> Allows users to select the ranges in case the table columns will change in the future
    Set flaggingList = Application.InputBox("Select your flagging list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not flaggingList Is Nothing Then
            If flaggingList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the flagging list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
    
    Set flaggingListResult = Application.InputBox("Select the column header cell in the flagging list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not flaggingListResult Is Nothing Then
            If flaggingListResult.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple cells selected! Please pick only the header cell in the flagging list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
    Set AllSizesOnHandList = Application.InputBox("Select your All On Hand list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not AllSizesOnHandList Is Nothing Then
            If AllSizesOnHandList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the all stock on hand list sheet and retry!", vbCritical
                Exit Sub
            End If
        End If

     flaggingListHeaderRowsCount = flaggingListResult.Row
     flaggingListHeaderRowNum = flaggingListResult.Row - 1
    
     AllSizesOnHandListHeaderRowNum = AllSizesOnHandList.Row - 1

    Application.ScreenUpdating = False
       
    '---> Allows users to compare the inventory list to the scan list in order to find matches
    For Each flaggingListStyleFabricColorCell In flaggingList
        flaggingListStyleFabricColorCriteria = Trim(flaggingListStyleFabricColorCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
     With AllSizesOnHandList
            Set foundMatchingAllSizesOnHandListStyleFabricColor = .Find(What:=flaggingListStyleFabricColorCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
     End With
    
    If foundMatchingAllSizesOnHandListStyleFabricColor Is Nothing Then
        If flaggingListStyleFabricColorCell.Row = flaggingListHeaderRowsCount Then
            If flaggingListResult.Cells(flaggingListStyleFabricColorCell.Row - flaggingListHeaderRowNum).value = vbNullString Then
                flaggingListResult.Cells(flaggingListStyleFabricColorCell.Row - flaggingListHeaderRowNum).value = "JDA Qty [Qty(Size) / Total Qty]"
                flaggingListResult.Font.FontStyle = "Bold"
            End If
            
        ElseIf flaggingListStyleFabricColorCell.Row > flaggingListHeaderRowsCount Then
               flaggingListResult.Cells(flaggingListStyleFabricColorCell.Row - flaggingListHeaderRowNum).value = "No Stock on hand"
        End If
      Else
        
    With AllSizesOnHandList
         currentAllSizesOnHandListCellRow = .Find(What:=flaggingListStyleFabricColorCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
    End With
        
        flaggingListResult.Cells(flaggingListStyleFabricColorCell.Row - flaggingListHeaderRowNum).value = AllSizesOnHandList.Cells(currentAllSizesOnHandListCellRow - AllSizesOnHandListHeaderRowNum, 3).value _
                                                                        & " / " & AllSizesOnHandList.Cells(currentAllSizesOnHandListCellRow - AllSizesOnHandListHeaderRowNum, 4).value & " units"
        

    End If
    Next flaggingListStyleFabricColorCell
        
        With flaggingListResult
             .EntireColumn.WrapText = True
             .ColumnWidth = 24
        End With
        Application.ScreenUpdating = True
    
    MsgBox "Process completed!"
    
Exit Sub
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
