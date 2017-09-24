Attribute VB_Name = "ULPCase"
Sub AddToCellMenu()
    Dim ContextMenu As CommandBar
    Dim MySubMenu As CommandBarControl

    ' Delete the controls first to avoid duplicates.
    Call DeleteFromCellMenu

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Add one built-in button(Save = 3) to the Cell context menu.
    ContextMenu.Controls.Add Type:=msoControlButton, ID:=3, before:=1

    ' Add one custom button to the Cell context menu.
    With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "ToggleCaseMacro"
        .FaceId = 59
        .Caption = "Toggle Case Upper/Lower/Proper"
        .Tag = "My_Cell_Control_Tag"
    End With

    ' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)

    With MySubMenu
        .Caption = "Case Menu"
        .Tag = "My_Cell_Control_Tag"

        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "UpperMacro"
            .FaceId = 100
            .Caption = "Upper Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "LowerMacro"
            .FaceId = 91
            .Caption = "Lower Case"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "ProperMacro"
            .FaceId = 95
            .Caption = "Proper Case"
        End With
    End With

    ' Add a separator to the Cell context menu.
    ContextMenu.Controls(4).BeginGroup = True
End Sub

Sub DeleteFromCellMenu()
    Dim ContextMenu As CommandBar
    Dim ctrl As CommandBarControl

    ' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")

    ' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "My_Cell_Control_Tag" Then
            ctrl.Delete
        End If
    Next ctrl

    ' Delete the custom built-in Save button.
    On Error Resume Next
    ContextMenu.FindControl(ID:=3).Delete
    On Error GoTo 0
End Sub

Sub ToggleCaseMacro()
    Dim selectedRange As Range
    Dim cell As Range

    On Error Resume Next
    Set selectedRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    For Each cell In selectedRange.Cells
        Select Case cell.value
        Case UCase(cell.value): cell.value = LCase(cell.value)
        Case LCase(cell.value): cell.value = StrConv(cell.value, vbProperCase)
        Case Else: cell.value = UCase(cell.value)
        End Select
    Next cell

    Application.ScreenUpdating = True
    
End Sub

Sub UpperMacro()
    Dim selectedRange As Range
    Dim cell As Range

    On Error Resume Next
    Set selectedRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Sub

Application.ScreenUpdating = False

    For Each cell In selectedRange.Cells
        cell.value = UCase(cell.value)
    Next cell

Application.ScreenUpdating = True

End Sub

Sub LowerMacro()
    Dim selectedRange As Range
    Dim cell As Range

    On Error Resume Next
    Set selectedRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Sub

Application.ScreenUpdating = False

    For Each cell In selectedRange.Cells
        cell.value = LCase(cell.value)
    Next cell

Application.ScreenUpdating = True

End Sub

Sub ProperMacro()
    Dim selectedRange As Range
    Dim cell As Range

    On Error Resume Next
    Set selectedRange = Intersect(Selection, _
        Selection.Cells.SpecialCells(xlCellTypeConstants, xlTextValues))
    On Error GoTo 0
    If selectedRange Is Nothing Then Exit Sub

Application.ScreenUpdating = False

    For Each cell In selectedRange.Cells
        cell.value = StrConv(cell.value, vbProperCase)
    Next cell

Application.ScreenUpdating = True

End Sub



