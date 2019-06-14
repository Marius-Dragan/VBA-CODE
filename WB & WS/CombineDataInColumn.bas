Attribute VB_Name = "ConsolidatingDatainColumn"
Option Explicit
Sub ConsolidatingDatainColumn()
Dim ws As Worksheet
Dim i As Long
Dim LRow As Long
Dim delRange As Range

Set ws = ActiveSheet

With ws
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        'Using a reverse loop to check for duplicates and write the result if matched on the cell above
        For i = LRow To 2 Step -1
             If .Range("A" & i).value = .Range("A" & i - 1).value Then
                .Range("B" & i - 1).value = .Range("B" & i - 1).value & ", " & .Range("B" & i).value
                .Range("A" & i).ClearContents
            End If
        Next i
        
        'Check for empty cells and delete
        For i = 6 To LRow
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
End With
End Sub
