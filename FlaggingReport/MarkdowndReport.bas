Attribute VB_Name = "MarkdownReport"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub EditMarkdownReport()

Dim WS As Worksheet
Dim delRange As Range
Dim lrow As Long, i As Long
Dim departmentDescription As String
Dim saveToFolderPath As String
Dim questionBoxPopUp As VbMsgBoxResult
Dim currentProgressBar As New ProgressDialogue

questionBoxPopUp = MsgBox("Are you sure you want to edit the markdown report worksheet?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit markdown report")
If questionBoxPopUp = vbNo Then Exit Sub

On Error GoTo ErrorHandler

saveToFolderPath = Environ("UserProfile") & "\Desktop"

Call CopySheet

Set WS = ActiveSheet

Columns("A:A").EntireColumn.Select
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(6, 1), Array(17, 1), Array(34, 1), Array(55, 1), _
        Array(69, 1), Array(92, 1), Array(100, 1), Array(107, 1), Array(122, 1), Array(132, 1), _
        Array(148, 1), Array(164, 1), Array(173, 1), Array(186, 1)), TrailingMinusNumbers:= _
        True
        ActiveCell.Select
    
Application.ScreenUpdating = False

With WS
       lrow = .Range("A" & .Rows.Count).End(xlUp).Row
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        currentProgressBar.Show
        
        For i = 5 To lrow
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Looping and deleting all rows where cell I is empty or has specific pattern of occurrences and inserting various formulas converting them to values " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:
            
            '--> Using upper case to check and write the department description for every cell otherwise the method will fail if cell values are lower case
            If UCase(.Range("D" & i).value) = UCase("WOMENS RTW") _
            Or UCase(.Range("D" & i).value) = UCase("MEN RTW") _
            Or UCase(.Range("D" & i).value) = UCase("DENIM") _
            Or UCase(.Range("D" & i).value) = UCase("LINGERIE") _
            Or UCase(.Range("D" & i).value) = UCase("SWIMWEAR") _
            Or UCase(.Range("D" & i).value) = UCase("MEN SWIMWEAR") _
            Or UCase(.Range("D" & i).value) = UCase("WRTW ACCESSORIES") _
            Or UCase(.Range("D" & i).value) = UCase("MEN BELTS") _
            Or UCase(.Range("D" & i).value) = UCase("BELTS") _
            Or UCase(.Range("D" & i).value) = UCase("HANDBAGS") _
            Or UCase(.Range("D" & i).value) = UCase("MEN HANDBAGS") _
            Or UCase(.Range("D" & i).value) = UCase("TRAVEL") _
            Or UCase(.Range("D" & i).value) = UCase("SMALL NON LEATHER GO") _
            Or UCase(.Range("D" & i).value) = UCase("MEN EYEWEAR") _
            Or UCase(.Range("D" & i).value) = UCase("EYEWEAR") _
            Or UCase(.Range("D" & i).value) = UCase("MENS SHOES") _
            Or UCase(.Range("D" & i).value) = UCase("WOMENS SHOES") _
            Or UCase(.Range("D" & i).value) = UCase("FRAGRANCES") _
            Or UCase(.Range("D" & i).value) = UCase("MEN SNLG") _
            Or UCase(.Range("D" & i).value) = UCase("SKINCARE") _
            Or UCase(.Range("D" & i).value) = UCase("ACCESSORIES") _
            Or UCase(.Range("D" & i).value) = UCase("MEN OTHER ACCS") _
            Or UCase(.Range("D" & i).value) = UCase("JEWELLERY") _
            Or UCase(.Range("D" & i).value) = UCase("KIDS") Or UCase(.Range("D" & i).value) = UCase("KIDS HANDBAGS") _
            Or UCase(.Range("D" & i).value) = UCase("ADIDAS") Or UCase(.Range("D" & i).value) = UCase("BOOKS") Then
                
                departmentDescription = .Range("D" & i).value
            Else
                .Range("D" & i).value = ""
            End If
            
            '--> Adding the department description the the empty cells
            If .Range("D" & i).value = "" Then
                .Range("D" & i).value = departmentDescription
            End If
            
            '--> Concatonatinc for full style/fabric/color
            If .Range("E" & i).value <> "" Then
            .Range("E" & i).value = .Range("E" & i).value & .Range("G" & i).value
            End If
            
            '--> Deleting empty rows in column I
            If Len(Trim(.Range("I" & i).value)) = 0 _
            Or .Range("I" & i).value = "AIL" _
            Or .Range("I" & i).value = "===============" _
            Or .Range("I" & i).value = "SKU" _
            Or .Range("I" & i).value = "---------------" Then
            
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
        lrow = .Range("I" & .Rows.Count).End(xlUp).Row
        
        '--> Inserting formula into the empty cells to copy data and convert them to values
        .Range("E5:E" & lrow - 1).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
        .Range("E5:E" & lrow - 1).value = .Range("E5:E" & lrow - 1).value
        
        '--> Inserting the formula in the entire column J5 in one go concatonating qty with size and converting it to values also clearing column H
        .Range("I5:I" & lrow - 1).Formula = "=J5 & ""("" & H5 & "")"""
        .Range("I5:I" & lrow - 1).value = .Range("I5:I" & lrow - 1).value
        .Range("H5:H" & lrow).Clear
        
        currentProgressBar.Hide
        currentProgressBar.Show
        
        
        '--> Using a reverse loop to append the value to the above cell
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        For i = lrow To 5 Step -1
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Using a reverse loop to append values from bottom row to the row above " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:
            
            If .Range("E" & i).value = .Range("E" & i - 1).value Then
                    .Range("I" & i - 1).value = .Range("I" & i - 1).value & ", " + .Range("I" & i).value
                    .Range("J" & i - 1).value = .Range("J" & i - 1).value + .Range("J" & i).value
                    .Range("I" & i).ClearContents
                    .Range("J" & i).ClearContents
                End If
        Next i
        
        .Range("H5:H" & lrow - 1).Formula = "=I5"
        .Range("H5:H" & lrow - 1).value = .Range("H5:H" & lrow - 1).value
        
        currentProgressBar.Hide
        currentProgressBar.Show
        
        '--> Delete rows where Cell I is empty
        currentProgressBar.Configure "Editing..." & "Please wait!", "Gathering info", i, lrow, , True, True
        For i = 5 To lrow
        
            currentProgressBar.SetValue i
            currentProgressBar.SetStatus "Delete rows where the rows on cell I are empty " & i & " out of " & lrow & " rows done"
            If currentProgressBar.cancelIsPressed Then GoTo CanceledBtnPressed:
        
            If Len(Trim(.Range("I" & i).value)) = 0 Then
                If delRange Is Nothing Then
                    Set delRange = .Rows(i)
                Else
                    Set delRange = Union(delRange, .Rows(i))
                End If
            End If
        Next i
        
        If Not delRange Is Nothing Then delRange.Delete
        
        Set delRange = Nothing
        
        .Columns.EntireColumn.AutoFit
        
   End With
   
   Call EditTable
   
   Call EditPrintProperties(WS)
   
   Call SaveAsToFolderPath

ScreenUpdate:
    Application.ScreenUpdating = True
    Unload currentProgressBar
    MsgBox "Process completed!" & vbNewLine & vbNewLine & "File path if save is successful:" & vbNewLine & saveToFolderPath, vbInformation, Title:="Markdown Detailes Report"
  
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
    Dim i As Integer
    
    MySheetName = "Edited_Markdown_Report"
    ActiveSheet.Name = "Raw_Data"

        If sheetExists("Edited_Markdown_Report") Then
            For i = 1 To Worksheets.Count
                If Worksheets(i).Name Like "*Raw_Data*" Then
            
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
Dim dateFormat As String
Dim storeCode As String
Dim lrow As Long

Set WS = ActiveSheet
dateFormat = Format(Now, " dd.mm.yyyy")
storeCode = Range("A5").value

With WS
    Rows(4).Delete
    Rows(2).Delete
    Rows(1).Delete
    .Range("B1").value = "City"
    .Range("C1").value = "Store name"
    .Range("D1").value = "Departments"
    .Range("E1").value = "Style/Fabric/Color"
    .Range("F1").value = "Description"
    .Columns(9).Delete
    .Columns(7).Delete
    .Columns(3).Delete
    .Columns(2).Delete
    .Columns(1).Delete
    lrow = .Range("A" & .Rows.Count).End(xlUp).Row
    
    .Range("A" & lrow).value = ""
    .Range("A" & lrow, "D" & lrow).Merge
    .Range("A" & lrow, "D" & lrow).value = "Grand Total"
    .Range("A" & lrow, "E" & lrow).Font.FontStyle = "Bold"
    .Range("A" & lrow, "E" & lrow).Font.size = 14
    .Range("A" & lrow, "D" & lrow).HorizontalAlignment = xlLeft
    .Range("E" & lrow).Formula = "=SUM(E2" & ":E" & lrow - 1 & ")"
    
    Call CreateTable
    
    .Rows(1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    .Range("A1:J1").Merge
    .Range("A1:J1").value = storeCode & " Markdown Detailed Report" & dateFormat
    .Range("A1:J1").Font.FontStyle = "Bold"
    .Range("A1:J1").Font.size = 26
    .Range("A1:J1").HorizontalAlignment = xlCenter
    .Range("A2").CurrentRegion.BorderAround xlContinuous, xlThin

End With
    

End Sub
Private Sub CreateTable()
    Dim lo As ListObject
    
    If Not TableExistsOnSheet(ActiveSheet, ActiveSheet.Name) Then
        Set lo = ActiveSheet.ListObjects.Add(xlSrcRange, [A2].CurrentRegion, , xlYes)
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
 
    lastRow = Range("E" & Rows.Count).End(xlUp).Row
    ActiveSheet.PageSetup.PrintArea = "A1:J" & lastRow
    
End Sub
Private Sub SaveAsToFolderPath()
'Set reference to Microsoft Scripting RunTime to see the properties and methods available in the IntelliSense
'The below 2 example will display the IntelliSense if the reference is set on
'Dim fso As Scripting.FileSystemObject
'Set fso = New Scripting.FileSystemObject

    Dim myFileName As String
    Dim newFolderPath As String
    Dim saveDetails As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Change file path to where you want to save the file
     newFolderPath = Environ("UserProfile") & "\Desktop\"
    
        If Not fso.FolderExists(newFolderPath) Then
               fso.CreateFolder newFolderPath
        End If
        
              myFileName = Range("A1").value
              saveDetails = newFolderPath & myFileName & ".xlsx"
              
        If Not fso.FileExists(saveDetails) Then
              
           If Not ActiveWorkbook.Saved Then
               ActiveWorkbook.SaveAs saveDetails, xlWorkbookDefault
           End If
        End If
        
    Set fso = Nothing
End Sub
