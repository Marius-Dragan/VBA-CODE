Attribute VB_Name = "RTV"
Option Explicit

Sub EditRTVfile()
'
'

    Dim ws As Worksheet
    Dim delRange As Range
    Dim lRow As Long, i As Long
    Dim questionBoxPopUp As VbMsgBoxResult
    Dim currentProgressBar As New ProgressDialogue

    questionBoxPopUp = MsgBox("Are you sure you want to edit the RTV worksheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit spot check file?")
    If questionBoxPopUp = vbNo Then Exit Sub

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    EditTable

    Set ws = ActiveSheet

 With ws
        lRow = .Range("A" & .Rows.Count).End(xlUp).Row

             '--> Insert a new column B and optionally L
        .Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        '.Columns(12).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

        '--> Inserting the formula in the entire column in one go and converting it to values
        .Range("B2:B" & lRow).Formula = "=A2 & "" ("" & J2 & "")"""
        '.Range("L2:L" & lRow).Formula = "=K2 & ""("" & J2 & "")"""
        .Range("B2:B" & lRow).value = .Range("B2:B" & lRow).value
        '.Range("L2:L" & lRow).value = .Range("L2:L" & lRow).value
        .Range("A2:A" & lRow).WrapText = True
        .Range("A2:A" & lRow).VerticalAlignment = xlTop
        .Range("A1:M1").VerticalAlignment = xlCenter
        .Range("A1:M1").HorizontalAlignment = xlCenter
        .Range("C2:L" & lRow).HorizontalAlignment = xlCenter
        .Range("C2:L" & lRow).VerticalAlignment = xlCenter

        '--> Copy the header from Col A to Col L so that we can delete the
        '--> Column G as it is not required anymore
        .Range("B1").value = "Carton number with qty scanned"
        .Range("J1").value = "QTY Scanned"
        '.Columns(11).Delete
        .Columns(1).Delete


      For i = lRow To 2 Step -1

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

        Set delRange = Nothing

                '--> Delete rows where Cell A is empty
        For i = 2 To lRow

            If Len(Trim(.Range("A" & i).value)) = 0 Then
                If delRange Is Nothing Then
                    Set delRange = .Rows(i)
                Else
                    Set delRange = Union(delRange, .Rows(i))
                End If
            End If
        Next i

        If Not delRange Is Nothing Then delRange.Delete
        
     
     .Cells.Rows.AutoFit
     .Cells.Columns.AutoFit

        End With

        editPrintProperties ws
        Application.ScreenUpdating = True
        MsgBox "Process completed!"
    
     Exit Sub
ErrorHandler:
Application.ScreenUpdating = True

    Debug.Print Err.Number; Err.Description
        MsgBox "Sorry, an error occured." & vbNewLine & vbNewLine & "Please print screen with the error message together with step by step commands that triggered the error to the developer in order to fix it." & vbNewLine & vbCrLf & Err.Number & " " & Err.Description, vbCritical, "Error!"
        'Resume ScreenUpdate
End Sub
Private Sub EditTable()

Dim ws As Worksheet
Dim headerRng As Range
Dim columnsToDelete As Range
Dim lastRow As Long
Dim lRow As Long, i As Long
Dim delRange As Range
Dim allBorders As Range

Set ws = ActiveSheet
Set headerRng = Range("A1", "W16")

With ws
       .Columns("A:W").UnMerge
        headerRng.Delete
    Set columnsToDelete = Application.Union(.Columns("A"), _
                                            .Columns("H:I"), _
                                            .Columns("K"), _
                                            .Columns("M:W"))
        columnsToDelete.Delete

        lRow = .Range("G" & .Rows.Count).End(xlUp).Row

         ActiveWindow.FreezePanes = False
        .Range("A2:A" & lRow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("A2:A" & lRow).value = .Range("A2:A" & lRow).value


        '--> Delete All rows where Cell B or Cell C are empty
        For i = 2 To lRow

            If Len(Trim(.Range("B" & i).value)) = 0 Or Len(Trim(.Range("C" & i).value)) = 0 Then
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

        Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("C1").value = "9 DIGIT SKU"
        .Range("C2:C" & lRow).Formula = "=Left(B2,9)"
        .Range("C2:C" & lRow).value = .Range("C2:C" & lRow).value

        With .Range("J1")
        .value = "Inventory List"
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

        .Range("K2:K" & lRow).Formula = "=I2-J2"
        .Range("K2:K" & lRow).NumberFormat = "0"
        .Range("A1:L1").Interior.Color = RGB(87, 175, 255)
        .Range("A1").CurrentRegion.Font.Size = 10
        .Range("A1").CurrentRegion.Font.Name = "Arial"
        .Range("A1").CurrentRegion.VerticalAlignment = xlTop
        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit
        .Range("A1").Activate
         Selection.AutoFilter

        lRow = Cells(Rows.Count, 3).End(xlUp).Row
        Range("A1:L" & lRow).Sort key1:=Range("C2:C" & lRow), _
        order1:=xlAscending, Orientation:=xlTopToBottom, Header:=xlYes

   End With

   Set allBorders = Range("A1").CurrentRegion

   With allBorders.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With

End Sub

Private Sub editPrintProperties(ws As Worksheet)

Dim lastRow As Long

     With ws.PageSetup
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
Sub RTV_ComparingTwoLists()

Dim ws As Worksheet

Dim inventoryListSkuCriteria As Variant
Dim scannerListSkuCriteria As Variant

Dim currentInventoryListCellRow As Long
Dim inventoryList As Range
Dim inventoryListResult As Range
Dim inventoryListSkuCell As Range
Dim foundMatchingInventoryListSku As Range

Dim currentScannerListCellRow As Long
Dim scannerList As Range
Dim scannerListResult As Range
Dim scannerListSkuCell As Range
Dim foundMatchingScannerListSku As Range

On Error GoTo ErrorHandler

'---> this methods needs all data to be visible in order to loop through all cells
For Each ws In ActiveWorkbook.Worksheets
     If ws.AutoFilterMode Then
         ws.AutoFilterMode = False
     End If
Next ws

'---> Allows users to select the ranges in case the table columns will change in the future
Set inventoryList = Application.InputBox("Select your inventory list range including header:", Type:=8)
    If Not inventoryList Is Nothing Then
        If inventoryList.Columns.Count = 1 Then
            Else
             MsgBox "Multiple columns selected! Please pick only one column in the inventory sheet and retry.", vbInformation
            Exit Sub
        End If
    End If

Set inventoryListResult = Application.InputBox("Select the column header cell in the invenotry list where to write the result:", Type:=8)
    If Not inventoryListResult Is Nothing Then
        If inventoryListResult.Rows.Count = 1 Then
            Else
             MsgBox "Multiple cells selected! Please pick only the header cell in the inventory sheet and retry.", vbInformation
            Exit Sub
        End If
    End If
    
Set scannerList = Application.InputBox("Select your scanner list range including header:", Type:=8)
    If Not scannerList Is Nothing Then
        If scannerList.Columns.Count = 1 Then
            Else
             MsgBox "Multiple columns selected! Please pick only one column in the inventory sheet and retry!", vbInformation
            Exit Sub
        End If
    End If
            
Set scannerListResult = Application.InputBox("Select the column header cell in the scanner list where to write the result:", Type:=8)
    If Not scannerListResult Is Nothing Then
        If scannerListResult.Rows.Count = 1 Then
            Else
             MsgBox "Multiple cells selected! Please pick only the header cell in the inventory sheet and retry!", vbInformation
            Exit Sub
        End If
    End If

'---> Allows users to compare the scan list to the inventory list in order to find matches
For Each scannerListSkuCell In scannerList
    scannerListSkuCriteria = scannerListSkuCell.value

    With inventoryList
        Set foundMatchingInventoryListSku = .Find(What:=scannerListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
    End With
       
If foundMatchingInventoryListSku Is Nothing Then
    If scannerListSkuCell.Row > 1 Then
        scannerListResult.Cells(scannerListSkuCell.Row).value = "Item not originally requested"
    End If
    Else

With inventoryList
        currentInventoryListCellRow = .Find(What:=scannerListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
    End With
     scannerListResult.Cells(scannerListSkuCell.Row).value = inventoryList.Cells(currentInventoryListCellRow, 3).value
End If
 
Next scannerListSkuCell

'---> Allows users to compare the inventory list to the scan list in order to find matches
For Each inventoryListSkuCell In inventoryList
    inventoryListSkuCriteria = inventoryListSkuCell.value

 With scannerList
        Set foundMatchingScannerListSku = .Find(What:=inventoryListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
    End With

If foundMatchingScannerListSku Is Nothing Then
    If inventoryListSkuCell.Row = 1 Then
        inventoryListResult.Cells(inventoryListSkuCell.Row).value = "QTY Scanned"
    End If
    If inventoryListSkuCell.Row > 1 Then
        inventoryListResult.Cells(inventoryListSkuCell.Row).value = "0"
    End If
    Else
    
With scannerList
        currentScannerListCellRow = .Find(What:=inventoryListSkuCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
    End With
    
    inventoryListResult.Cells(inventoryListSkuCell.Row).value = scannerList.Cells(currentScannerListCellRow, 7).value

End If
Next inventoryListSkuCell

'---> apply the filters back to all worksheets
For Each ws In ActiveWorkbook.Worksheets
         ws.Range("A1").AutoFilter
Next ws

MsgBox "Process completed!"

ErrorHandler:
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
