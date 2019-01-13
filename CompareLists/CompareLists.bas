Attribute VB_Name = "ComparingLists"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.

Sub ComparingTwoLists()
'--> Working version including to pick what column you what data
'--> Need to add dependency SmartUtlilities

    'Source list
    Dim cellCriteria As Variant
    Dim sourceList As Range
    Dim sourceListResult As Range
    Dim sourceListCell As Range
    Dim sourceListHeaderRowsCount As Integer
    Dim sourceListHeaderRowNum As Integer
    
    
    'Comparing against list
    Dim compareAgainstListCellRowNum As Long
    Dim compareAgainstList As Range
    Dim desiredResultColumn As Range
    Dim compareAgainstListHeaderRowNum As Integer
    Dim compareAgainstListColumnsCount As Integer
    
    'Variable to hold if match found on the comparingAgainstList
    Dim foundMatch As Range
    Dim firstFoundMatchingAddress As String
    
    On Error GoTo ErrorHandler
    
    Call SmartUtilities.ResetFilters
        
    '---> Allows users to select the ranges in case the table columns will change in the future
    Set sourceList = Application.InputBox("Select your source list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not sourceList Is Nothing Then
            If sourceList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the source list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
    
    Set sourceListResult = Application.InputBox("Select the column header cell in the source list where to write the result:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not sourceListResult Is Nothing Then
            If sourceListResult.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple cells selected! Please pick only the header cell in the source list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
    Set compareAgainstList = Application.InputBox("Select your comparing list range including header:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not compareAgainstList Is Nothing Then
            If compareAgainstList.Columns.Count = 1 Then
                Else
                 MsgBox "Multiple columns selected! Please pick only one column in the comparing list sheet and retry.", vbCritical
                Exit Sub
            End If
        End If
        
    Set desiredResultColumn = Application.InputBox("Select the column header cell in the comparing list that has the data needed:", Default:="'" & ActiveSheet.Name & "'!", Type:=8)
        If Not compareAgainstList Is Nothing Then
            If compareAgainstList.Columns.Count = 1 Or compareAgainstList.Rows.Count = 1 Then
                Else
                 MsgBox "Multiple columns or rows selected! Please pick only the header cell in the comparing list sheet that has the needed data and retry.", vbCritical
                Exit Sub
            End If
        End If
    
    sourceListHeaderRowsCount = sourceListResult.Row
    sourceListHeaderRowNum = sourceListResult.Row - 1
    
    
     compareAgainstListHeaderRowNum = desiredResultColumn.Row - 1
     compareAgainstListColumnsCount = (desiredResultColumn.Column - compareAgainstList.Column) + 1
    
    Application.ScreenUpdating = False
       
    '---> Allows users to compare the source list to the comparing against list in order to find matches
    For Each sourceListCell In sourceList
        cellCriteria = Trim(sourceListCell.value) 'Using trim to delete the extra space from the data otherwhise it will throw an error
    
     With compareAgainstList
            Set foundMatch = .Find(What:=cellCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) 'finds a match
     End With
    
    If foundMatch Is Nothing Then
        If sourceListCell.Row = sourceListHeaderRowsCount Then
            If sourceListResult.Cells(sourceListCell.Row - sourceListHeaderRowNum).value = vbNullString Then
                sourceListResult.Cells(sourceListCell.Row - sourceListHeaderRowNum).value = "Comparing List"
                sourceListResult.Font.FontStyle = "Bold"
            End If
            
        ElseIf sourceListCell.Row > sourceListHeaderRowsCount Then
               sourceListResult.Cells(sourceListCell.Row - sourceListHeaderRowNum).value = "No match found"
        End If
      Else
        
       firstFoundMatchingAddress = foundMatch.Address
       
    Do
        
    With compareAgainstList
         'compareAgainstListCellRowNum = .Find(What:=cellCriteria, After:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Row
          compareAgainstListCellRowNum = foundMatch.Row
    End With
        
        sourceListResult.Cells(sourceListCell.Row - sourceListHeaderRowNum).value = compareAgainstList.Cells(compareAgainstListCellRowNum - compareAgainstListHeaderRowNum, compareAgainstListColumnsCount).value
    
    Loop While foundMatch.Address <> firstFoundMatchingAddress
    
     End If
     
    Next sourceListCell
        
        With sourceListResult
             .ColumnWidth = 24
             .EntireColumn.WrapText = True
             .EntireColumn.AutoFit
             .EntireRow.AutoFit
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



