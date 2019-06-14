Attribute VB_Name = "FillInDataSets"
Option Explicit
Sub FillInDataCSV()
Dim ws As Worksheet
Dim LRow As Long
Dim i As Long
Dim cell As Range

Set ws = ActiveSheet

    With ws
        LRow = .Range("AM" & .Rows.Count).End(xlUp).Row
        
        For i = 5 To LRow
        
        If i = 5 And Len(Trim(.Range("AN" & i).value)) = 0 Then
            .Range("AN" & i).value = 1
        End If
        If i = 5 And Len(Trim(.Range("BL" & i).value)) = 0 Then
            .Range("BL" & i).value = 1
        End If
        
                If Len(Trim(.Range("X" & i).value)) = 0 Then
                    .Range("X" & i).value = .Range("X" & i - 1).value
                End If
                If Len(Trim(.Range("Y" & i).value)) = 0 Then
                    .Range("Y" & i).value = .Range("Y" & i - 1).value
                End If
                If Len(Trim(.Range("AN" & i).value)) = 0 Then
                    .Range("AN" & i).value = .Range("AN" & i - 1).value + 1
                End If
                If Len(Trim(.Range("AR" & i).value)) = 0 Then
                    .Range("AR" & i).value = .Range("AR" & i - 1).value
                End If
                If Len(Trim(.Range("AV" & i).value)) = 0 Then
                    .Range("AV" & i).value = .Range("AV" & i - 1).value
                End If
                If Len(Trim(.Range("AW" & i).value)) = 0 Then
                    .Range("AW" & i).value = .Range("AW" & i - 1).value
                End If
                If Len(Trim(.Range("BL" & i).value)) = 0 Then
                    .Range("BL" & i).value = .Range("BL" & i - 1).value + 1
                End If
        Next i
    End With
End Sub


