Attribute VB_Name = "EditConsReport"
Sub EditConsReport()
Attribute EditConsReport.VB_ProcData.VB_Invoke_Func = " \n14"

Dim questionBoxPopUp As VbMsgBoxResult
Dim lrow As Long, i As Long
Dim ws As Worksheet

questionBoxPopUp = MsgBox("Are you sure you want to edit open consignment report?", vbQuestion + vbYesNo + vbDefaultButton1, "Edit consignment report")
    If questionBoxPopUp = vbNo Then Exit Sub

 Application.ScreenUpdating = False
 
    Call CopySheet
    
    Set ws = ActiveSheet
 
    Range("A8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A8"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(5, 1), Array(13, 1), Array(19, 1), Array(44, 4), _
        Array(52, 4), Array(62, 1), Array(78, 1), Array(81, 1), Array(101, 1), Array(110, 1), Array _
        (137, 1), Array(141, 1), Array(148, 1), Array(158, 1), Array(170, 1), Array(185, 1), Array( _
        200, 9)), TrailingMinusNumbers:=True
        
        changeProperties ws
        Call DeleteRowBasedOnCriteria
        Range("A1").EntireColumn.AutoFit
        
        With ws
        '--> Find the last row
        lrow = .Range("G" & .Rows.Count).End(xlUp).Row
        
        '--> Insert a new column between G and H
        .Columns(8).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
        '--> Insert a formula =G3 & "(" & M3 & ")" in H3
        '--> Inserting the formula in the entire column in one go and converting it to values
        .Range("H3:H" & lrow).Formula = "=G3 & M3"
        .Range("H3:H" & lrow).value = .Range("H3:H" & lrow).value
        '--> Enter the header in Col H so that we can delete the Column G as it is not required anymore
        .Range("H2").value = "Style/Fabric/Colour"
        .Columns(7).Delete
       
        End With
        
        Cells.EntireColumn.AutoFit
        Range("A2").Select
        Call SaveAsToFolderPath
        
        Application.ScreenUpdating = True
        
End Sub
Private Sub CopySheet()
 
    Dim MySheetName As String
    MySheetName = "Open Consignment Report"
    
        ActiveSheet.Name = "Raw Data"
        
        If sheetExists("Edited Spot Check") = True Then
            MsgBox "Sheet named " & "'Open Consignment Report'" & " already exists. Please rename if you need another copy!", vbInformation, "Sheet exists!"
    End
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
Sub DeleteRowBasedOnCriteria()
Dim RowToTest As Long
Dim ws As Worksheet
Dim lrow As Long, i As Long

Set ws = ActiveSheet

For RowToTest = Cells(Rows.Count, 6).End(xlUp).Offset(-2, 0).Row To 2 Step -1

With Cells(RowToTest, 6)
    If .value = "Tota" Then
    Rows(RowToTest).EntireRow.Delete
    End If
    End With
With Cells(RowToTest, 6)
    If .value = "--------" Then
     Rows(RowToTest).EntireRow.Delete
     End If
     End With
With Cells(RowToTest, 6)
    If .value = "----------" Then
     Rows(RowToTest).EntireRow.Delete
     End If
     End With
With Cells(RowToTest, 6)
    If .value = "" Then
     Rows(RowToTest).EntireRow.Delete
     End If
     End With

Next RowToTest

 

    With ws
        lrow = .Range("F" & .Rows.Count).End(xlUp).Row
        For i = 3 To lrow
        With Cells(i, 6)
        If .value = "ReturnBy" Then
     Rows(i).EntireRow.Delete
     End If
   End With
     Next i
       End With
End Sub
Private Sub changeProperties(ws As Worksheet)

Dim tableRng As Range
Dim columnsToHide As Range
Dim headerFormat As Range

Set headerFormat = Range(Cells(8, 1), Cells(8, Columns.Count).End(xlToLeft)).Columns
Set tableRng = Range("A8").CurrentRegion
Set ws = ActiveSheet

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
    
        With ws.Range("A1:Q1")
        .Merge
        .value = "Stella McCartney UK - Open Consignment Report"
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Color = vbBlack
        .Font.size = 26
        .HorizontalAlignment = xlCenter
        End With
        

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
     newFolderPath = Environ("UserProfile") & "\Desktop\Marius\"
     
        If Not fso.FolderExists(newFolderPath) Then
               fso.CreateFolder newFolderPath
        End If
        
              dateFormat = Format(Now, " dd.mm.yyyy")
              'myFileName = Range("A1").value
              myFileName = Range("A3").value & Mid(Range("A1").value, 20, 50)
              saveDetails = newFolderPath & myFileName & dateFormat & ".xlsx"
              
        If Not fso.FileExists(saveDetails) Then
              
           If Not ActiveWorkbook.Saved Then
               ActiveWorkbook.SaveAs saveDetails, xlWorkbookDefault
           End If
        End If
        
    Set fso = Nothing
End Sub


