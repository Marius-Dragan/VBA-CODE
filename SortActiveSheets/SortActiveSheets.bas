Attribute VB_Name = "SortActiveSheets"
Sub SortActiveWorksheets()


Dim i As Integer
Dim j As Integer
Dim questionBoxPopUp As VbMsgBoxResult


'
' Prompt the user as which direction they wish to
' sort the worksheets.
'
   questionBoxPopUp = MsgBox("Sort Sheets in Ascending Order?" & Chr(10) _
     & "Clicking No will sort in Descending Order.", _
     vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
   For i = 1 To Sheets.Count
      For j = 1 To Sheets.Count - 1
'
' If the answer is Yes, then sort in ascending order.
'
         If questionBoxPopUp = vbYes Then
            If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
'
' If the answer is No, then sort in descending order.
'
         ElseIf questionBoxPopUp = vbNo Then
            If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
               Sheets(j).Move After:=Sheets(j + 1)
            End If
         End If
      Next j
   Next i
    MsgBox "Process completed!", vbInformation
End Sub



