Attribute VB_Name = "RTVReport"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub EditRTVReport()

    Dim WS As Worksheet
    Dim delRange As Range
    Dim basketIDRange As Variant
    Dim lrow As Long, i As Long
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim currentProgressBar As New ProgressDialogue

    questionBoxPopUp = MsgBox("Are you sure you want to edit the RTV report?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit RTV Report?")
    If questionBoxPopUp = vbNo Then Exit Sub
    
    Call CopySheet

    On Error GoTo ErrorHandler
    
    basketIDRange = Range("O1").value
    
    Application.ScreenUpdating = False

    Call EditTable

    Set WS = ActiveSheet

 With WS
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        currentProgressBar.Show


             '--> Insert a new column B and optionally L
        .Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        '.Columns(12).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

        '--> Inserting the formula in the entire column in one go and converting it to values
        .Range("B2:B" & lrow).Formula = "=A2 & "" ("" & J2 & "")"""
        '.Range("L2:L" & lRow).Formula = "=K2 & ""("" & J2 & "")"""
        .Range("B2:B" & lrow).value = .Range("B2:B" & lrow).value
        '.Range("L2:L" & lRow).value = .Range("L2:L" & lRow).value
        .Range("A2:A" & lrow).WrapText = True
        .Range("A2:A" & lrow).VerticalAlignment = xlTop
        .Range("A1:M1").VerticalAlignment = xlCenter
        .Range("A1:M1").HorizontalAlignment = xlCenter
        .Range("C2:L" & lrow).HorizontalAlignment = xlCenter
        .Range("C2:L" & lrow).VerticalAlignment = xlCenter

        '--> Copy the header from Col A to Col L so that we can delete the
        '--> Column G as it is not required anymore
        .Range("B1").value = "ID Basket: " & basketIDRange & " - Carton number with qty scanned"
        .Range("J1").value = "QTY Scanned"
        '.Columns(11).Delete
        .Columns(1).Delete


      For i = lrow To 2 Step -1
      
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Using a reverse loop to append values from bottom row to the row above for Column A and I " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

         '--> Using a reverse loop to append values from bottom row to the row above for Column A and I
        '--> After appending clear the cell A so that we can later delete the row
            If .Range("C" & i).value = .Range("C" & i - 1).value Then
                .Range("A" & i - 1).value = .Range("A" & i - 1).value & ", " & .Range("A" & i).value
                .Range("I" & i - 1).value = .Range("I" & i - 1).value + .Range("I" & i).value
                '---> Extra line in case of scanning with 2 scanners
                '.Range("J" & i - 1).value = .Range("J" & i - 1).value & ", " & .Range("J" & i).value
                .Range("A" & i).ClearContents
                '.Range("J" & i).ClearContents
            End If
        Next i
        
        currentProgressBar.Hide
        currentProgressBar.Show
        
        '--> Delete rows where Cell A is empty
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        For i = 2 To lrow
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Delete rows where Cell A is empty " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:

            If Len(Trim(.Range("A" & i).value)) = 0 Then
                If delRange Is Nothing Then
                    Set delRange = .Rows(i)
                Else
                    Set delRange = Union(delRange, .Rows(i))
                End If
            End If
        Next i

        If Not delRange Is Nothing Then delRange.Delete
        
        Set delRange = Nothing
     
     .Cells.Rows.AutoFit
     .Cells.Columns.AutoFit

        End With

        EditPrintProperties WS
        Application.ScreenUpdating = True
        Unload currentProgressBar
        MsgBox "Process completed!", vbInformation, Title:="RTV Report"
        
    Exit Sub
CanceledBtnPressed:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Cancelled By User.", vbInformation
    
     Exit Sub
ErrorHandler:
Application.ScreenUpdating = True

    Debug.Print Err.Number; Err.Description
        MsgBox "Sorry, an error occured." & vbNewLine & vbNewLine & "Please print screen with the error message together with step by step commands that triggered the error to the developer in order to fix it." & vbNewLine & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"
        'Resume ScreenUpdate
End Sub
Private Sub CopySheet()
 
    Dim MySheetName As String
    MySheetName = "Scanner"
    Dim i As Integer

        If sheetExists("Scanner") Then
            For i = 1 To Worksheets.Count
                If Worksheets(i).Name Like "*Scanner*" Then
            
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
Private Sub EditTable()

Dim WS As Worksheet
Dim headerRng As Range
Dim columnsToDelete As Range
Dim lastRow As Long
Dim lrow As Long, i As Long
Dim delRange As Range
Dim allBorders As Range

Set WS = ActiveSheet
Set headerRng = Range("A1", "W16")

With WS
       .Columns("A:W").UnMerge
        headerRng.Delete
    Set columnsToDelete = Application.Union(.Columns("A"), _
                                            .Columns("H:I"), _
                                            .Columns("K"), _
                                            .Columns("M:W"))
        columnsToDelete.Delete

        lrow = .Range("G" & .Rows.Count).End(xlUp).Row

         ActiveWindow.FreezePanes = False
         
         '--> Inserting furmula into the empty cells to copy data and convert them to values
        .Range("A2:A" & lrow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A2:A" & lrow).value = .Range("A2:A" & lrow).value


        '--> Delete All rows where Cell B or Cell C are empty
        For i = 2 To lrow

            If Len(Trim(.Range("B" & i).value)) = 0 Or Len(Trim(.Range("C" & i).value)) = 0 Then
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

        Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("C1").value = "9 DIGIT SKU"
        .Range("C2:C" & lrow).Formula = "=Left(B2,9)"
        .Range("C2:C" & lrow).value = .Range("C2:C" & lrow).value

        With .Range("J1")
        .value = "Inventory List (On Hand Qty)"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .BorderAround xlContinuous, xlThin
        End With

        With .Range("K1")
        .value = "Variance"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .BorderAround xlContinuous, xlThin
        End With

        With .Range("L1")
        .value = "Comments"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .BorderAround xlContinuous, xlThin
        End With

        .Range("K2:K" & lrow).Formula = "=I2-J2"
        .Range("K2:K" & lrow).NumberFormat = "0"
        .Range("A1:L1").Interior.Color = RGB(87, 175, 255)
        .Range("A1").CurrentRegion.Font.size = 10
        .Range("A1").CurrentRegion.Font.Name = "Arial"
        .Range("A1").CurrentRegion.VerticalAlignment = xlTop
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        .Range("A1").Activate
         Selection.AutoFilter

        lrow = Cells(Rows.Count, 3).End(xlUp).Row
        Range("A1:L" & lrow).Sort key1:=Range("C2:C" & lrow), _
        order1:=xlAscending, Orientation:=xlTopToBottom, Header:=xlYes

   End With

   Set allBorders = Range("A1").CurrentRegion

   With allBorders.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With

End Sub

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
    ActiveSheet.PageSetup.PrintArea = "A1:L" & lastRow

End Sub
Sub RTV_ComparingInventoryListToScannerList()
'--> Need to add dependency SmartUtlilities
'--> The comparing process takes into consideration that both lists contains unique values

    'Source list
    Dim currentInventoryListCellRow As Long
    Dim inventoryList As Range
    Dim inventoryListResult As Range
    Dim inventoryListSkuCell As Range
    Dim foundMatchinginventoryListSku As Range
    Dim inventoryListHeaderRowsCount As Integer
    Dim inventoryListHeaderRowNum As Integer
    
    'Comparing against list
    Dim currentScannerListCellRow As Long
    Dim scannerList As Range
    Dim scannerListResult As Range
    Dim scannerListSkuCell As Range
    Dim foundMatchingScannerListSku As Range
    Dim scannerListHeaderRowsCount As Integer
    Dim scannerListHeaderRowNum As Integer
    
    'Variables to hold if match found
    Dim inventoryListSkuCriteria As Variant
    Dim scannerListSkuCriteria As Variant
    
    
    On Error GoTo ErrorHandler
    
    '---> this method needs all data to be visible in order to loop through all cells
    Call SmartUtilities.ResetFilters
    
    '---> Allows users to select the ranges in case the table columns will change in the future
    Set inventoryList = Application.InputBox("Select your inventory list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not inventoryList Is Nothing Then
            If inventoryList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the inventory sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
    
    Set inventoryListResult = Application.InputBox("Select the column header cell in the invenotry list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not inventoryListResult Is Nothing Then
            If inventoryListResult.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple cells selected! Please pick only the header cell in the inventory sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
    Set scannerList = Application.InputBox("Select your scanner list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not scannerList Is Nothing Then
            If scannerList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the scanner sheet and retry!", vbCritical
                Exit Sub
            End If
        End If
                
    Set scannerListResult = Application.InputBox("Select the column header cell in the scanner list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not scannerListResult Is Nothing Then
            If scannerListResult.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple cells selected! Please pick only the header cell in the scanner sheet and retry!", vbCritical
                Exit Sub
            End If
        End If
        
        
     inventoryListHeaderRowsCount = inventoryListResult.Row
     inventoryListHeaderRowNum = inventoryListResult.Row - 1
    
    
     scannerListHeaderRowsCount = scannerListResult.Row
     scannerListHeaderRowNum = scannerListResult.Row - 1
    
    Application.ScreenUpdating = False
    
    '---> Allows users to compare the scan list to the inventory list in order to find matches
    For Each scannerListSkuCell In scannerList
        scannerListSkuCriteria = Trim(scannerListSkuCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
        With inventoryList 'If the column heading of both lists match the it will retrive the first row of heading
            Set foundMatchinginventoryListSku = .Find(What:=scannerListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
        End With
           
    If foundMatchinginventoryListSku Is Nothing Then
        If scannerListSkuCell.Row = scannerListHeaderRowsCount Then
            If scannerListResult.Cells(scannerListSkuCell.Row - scannerListHeaderRowNum).value = vbNullString Then
                scannerListResult.Cells(scannerListSkuCell.Row - scannerListHeaderRowNum).value = "Inventory List " & inventoryList.Cells(2, 4).value & " (On Hand Qty)"
                scannerListResult.Font.FontStyle = "Bold"
            End If
            
            ElseIf scannerListSkuCell.Row > scannerListHeaderRowsCount Then
                scannerListResult.Cells(scannerListSkuCell.Row - scannerListHeaderRowNum).value = "Item not originally requested"
            End If
      Else
    
    With inventoryList
            'Testing to see if same result needs to be done
            currentInventoryListCellRow = foundMatchinginventoryListSku.Row 'To delete this row and uncomment below if not working
            'currentInventoryListCellRow = .Find(What:=scannerListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
        End With
         scannerListResult.Cells(scannerListSkuCell.Row - scannerListHeaderRowNum).value = inventoryList.Cells(currentInventoryListCellRow - inventoryListHeaderRowNum, 3).value
    End If
     
    Next scannerListSkuCell
    
    '---> Allows users to compare the inventory list to the scan list in order to find matches
    For Each inventoryListSkuCell In inventoryList
        inventoryListSkuCriteria = Trim(inventoryListSkuCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
     With scannerList
            Set foundMatchingScannerListSku = .Find(What:=inventoryListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
     End With
    
    If foundMatchingScannerListSku Is Nothing Then
        If inventoryListSkuCell.Row = inventoryListHeaderRowsCount Then
            If inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = vbNullString Then
                inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = "QTY Scanned"
                inventoryListResult.Font.FontStyle = "Bold"
            End If
            
            ElseIf inventoryListSkuCell.Row > inventoryListHeaderRowsCount Then
                   inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = "0"
            End If
      Else
        
    With scannerList 'If the column heading of both lists match the it will retrive the first row of heading
         currentScannerListCellRow = foundMatchingScannerListSku.Row 'To delete this row and uncomment below if tests faild
         'currentScannerListCellRow = .Find(What:=inventoryListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
    End With
        
        inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = scannerList.Cells(currentScannerListCellRow - scannerListHeaderRowNum, 7).value
    
    End If
    Next inventoryListSkuCell
    
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
Sub RTV_ComparingInventoryListToConsList()
'--> Need to add dependency SmartUtlilities

    'Source list
    Dim inventoryList As Range
    Dim inventoryListResult As Range
    Dim inventoryListSkuCell As Range
    Dim inventoryListHeaderRowsCount As Integer
    Dim inventoryListHeaderRowNum As Integer
    Dim foundMatchinginventoryListSku As Range
    
    'Comparing against list
    Dim consList As Range
    Dim consListCellRow As Long
    Dim consListHeaderRowNum As Integer
    
    'Output variables
    Dim firstFoundMatchingAddress As String
    Dim countSkus As Long
    
    'Variable to hold if match found on the comparing against list
    Dim inventoryListSkuCriteria As Variant
    
     On Error GoTo ErrorHandler
    
    '---> this method needs all data to be visible in order to loop through all cells
    SmartUtilities.ResetFilters
    
    '---> Allows users to select the ranges in case the table columns will change in the future
    Set inventoryList = Application.InputBox("Select your all on hand list list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not inventoryList Is Nothing Then
            If inventoryList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the all on hand list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
     Set inventoryListResult = Application.InputBox("Select the column header cell in the all on hand list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not inventoryListResult Is Nothing Then
            If inventoryListResult.Rows.Count = 1 Then
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
        
    inventoryListHeaderRowsCount = inventoryListResult.Row
    inventoryListHeaderRowNum = inventoryListResult.Row - 1
    
    consListHeaderRowNum = consList.Row - 1
      
        Application.ScreenUpdating = False
        
     If consList.End(xlToRight).value <> "Matching skus" Then
        consList.End(xlToRight).Offset(0, 1).value = "Matching skus"
     End If
     
    '---> Allows users to compare the consignment list to the inventory list in order to find matches
    For Each inventoryListSkuCell In inventoryList
        inventoryListSkuCriteria = Trim(inventoryListSkuCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
        With consList
            Set foundMatchinginventoryListSku = .Find(What:=inventoryListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, _
                                                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, _
                                                SearchFormat:=False) 'finds a match
        End With
           
    If foundMatchinginventoryListSku Is Nothing Then
        If inventoryListSkuCell.Row = inventoryListHeaderRowsCount Then
            If inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = vbNullString Then
                inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = "Open Consignment List"
                inventoryListResult.Font.FontStyle = "Bold"
            End If
        End If
        
      Else
    
            consListCellRow = foundMatchinginventoryListSku.Row
            firstFoundMatchingAddress = foundMatchinginventoryListSku.Address
                                   
            Do '---> Looping through all instances of a value and write result
                  
                With consList
                    If inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = vbNullString Then
                    
                         inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = .Cells(consListCellRow - consListHeaderRowNum, 6).value & " on cons " _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, -7).value _
                                                                                        & " (" _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, -6).value _
                                                                                        & ") " _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, -5).value
                    Else
                        
                         inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value = inventoryListResult.Cells(inventoryListSkuCell.Row - inventoryListHeaderRowNum).value _
                                                                                        & " // " _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, 6).value & " on cons " _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, -7).value _
                                                                                        & " (" _
                                                                                        & .Cells(consListCellRow - consListHeaderRowNum, -6).value _
                                                                                        & ") " _
                                                                                       & .Cells(consListCellRow - consListHeaderRowNum, -5).value
                        
                    End If
                    
                    countSkus = countSkus + 1
                    
                    If foundMatchinginventoryListSku.End(xlToRight) <> "Found" Then
                        foundMatchinginventoryListSku.End(xlToRight).Offset(0, 1).value = "Found"
                    End If
                    
                    Set foundMatchinginventoryListSku = .FindNext(foundMatchinginventoryListSku)
                    
                End With
                
            Loop While Not foundMatchinginventoryListSku Is Nothing And foundMatchinginventoryListSku.Address <> firstFoundMatchingAddress
        
    End If
        
    Next inventoryListSkuCell
     
        Application.ScreenUpdating = True
        
        MsgBox "Process completed!" & vbNewLine & vbNewLine & "Found " & countSkus & " skus in the consignment list." & vbCrLf, vbInformation, "Comparing lists"

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

