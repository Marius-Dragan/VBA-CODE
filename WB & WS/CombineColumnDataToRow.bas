Attribute VB_Name = "CombineColumnDataToRow"
Option Explicit
Sub CombineColumnDataToRow()
'Consoilidate data in column to one cell in the worksheet separating each row in the column by ";"
Dim ws As Worksheet
Dim srcCell As Range
Dim destCell As Range

Set ws = ActiveSheet
Set destCell = ws.Range("D1")
Application.ScreenUpdating = False

    For Each srcCell In Selection
        'destCell.Activate
        destCell = destCell.value & ";" & srcCell.value
    Next srcCell
    
Application.ScreenUpdating = True

End Sub

