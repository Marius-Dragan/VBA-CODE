VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DryCleaningForm 
   Caption         =   "Dry Cleaning Form SMC"
   ClientHeight    =   12096
   ClientLeft      =   100
   ClientTop       =   460
   ClientWidth     =   23040
   OleObjectBlob   =   "DryCleaningForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "DryCleaningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SmcDryCleaningForm As String
Dim Dic As Object
Dim CorrectLoginDetails As Boolean
Dim UserList As Variant

Public Enum FormDataModes
    AddNewEmployeeBtn = 0
    RemoveEmployeeBtn = 1
    AddNewItemBtn = 2
    RemoveItemBtn = 3
    AdminModeBtn = 4
    AllBtnsEnabled = 5
End Enum

Public Enum FormDataEnableds
    enablegroup = 0
    DisableGroup = 1
End Enum

Private pFormDataEnabled As FormDataEnableds

Private pFormDataMode As FormDataModes
Public Property Let FormDataEnabled(Enabled As FormDataEnableds)
    pFormDataEnabled = Value
    
    If Enabled = enablegroup Then
          
    ElseIf Enabled = DisableGroup Then
        
    End If
End Property
Public Property Let FormDataMode(Value As FormDataModes)
    
    pFormDataMode = Value
    
    '--> Setting command button state
    If Value = AddNewEmployeeBtn Then
        'Me.AddNewEmployeeBtnPressed.Enabled = True
        Me.RemoveEmployeeBtnPressed.Enabled = False
        Me.RemoveEmployeeBtnPressed.Enabled = False
        Me.AddNewItemBtnPressed.Enabled = False
        Me.RemoveItemBtnPress.Enabled = False
        Me.AdminModeBtnPressed.Enabled = False
        
    ElseIf Value = RemoveEmployeeBtn Then
        Me.AddNewEmployeeBtnPressed.Enabled = False
        'Me.RemoveEmployeeBtnPressed.Enabled = False
        Me.AddNewItemBtnPressed.Enabled = False
        Me.RemoveItemBtnPress.Enabled = False
        Me.AdminModeBtnPressed.Enabled = False
        
    ElseIf Value = AddNewItemBtn Then
        Me.AddNewEmployeeBtnPressed.Enabled = False
        Me.RemoveEmployeeBtnPressed.Enabled = False
        'Me.AddNewItemBtnPressed.Enabled = False
        Me.RemoveItemBtnPress.Enabled = False
        Me.AdminModeBtnPressed.Enabled = False
        
    ElseIf Value = RemoveItemBtn Then
        Me.AddNewEmployeeBtnPressed.Enabled = False
        Me.RemoveEmployeeBtnPressed.Enabled = False
        Me.AddNewItemBtnPressed.Enabled = False
        'Me.RemoveItemBtnPress.Enabled = False
        Me.AdminModeBtnPressed.Enabled = False
        
    ElseIf Value = AdminModeBtn Then
        Me.AddNewEmployeeBtnPressed.Enabled = False
        Me.RemoveEmployeeBtnPressed.Enabled = False
        Me.AddNewItemBtnPressed.Enabled = False
        Me.RemoveItemBtnPress.Enabled = False
        'Me.AdminModeBtnPressed.Enabled = False
        
    ElseIf Value = AllBtnsEnabled Then
        Me.AddNewEmployeeBtnPressed.Enabled = True
        Me.RemoveEmployeeBtnPressed.Enabled = True
        Me.AddNewItemBtnPressed.Enabled = True
        Me.RemoveItemBtnPress.Enabled = True
        Me.AdminModeBtnPressed.Enabled = True
    End If
    
End Property


Private Sub EmployeeDropDownList_AfterUpdate()
        Call ActivateSheet(EmployeeDropDownList.Value)
        Call GenerateOrderNo
        Call UpdateEmployeeItemsToClean
        
        ActiveSheet.Range("B4").Value = Me.CurrentDate.Value
        
        Me.OrderLbl1.Enabled = True
        Me.Item1.Enabled = True
        Me.Qty1.Enabled = True
        Me.UnitPrice1.Enabled = True
        Me.SubTotal1.Enabled = True
        Me.OptionIn1.Enabled = True
        Me.OptionOut1.Enabled = True
        Me.Comments1.Enabled = True
        Me.ReturnDate1.Enabled = True
        
End Sub
Public Function CollectionDate(InputDate As Date) As Date
    
    Select Case Weekday(InputDate)
        Case vbTuesday
            CollectionDate = CDate(Date - Weekday(Date, vbSunday) + 4)
            'Debug.Print Weekday(Date, vbSunday)
        Case vbThursday
            CollectionDate = CDate(Date - Weekday(Date, vbSunday) + 6)
            'Debug.Print Weekday(Date, vbSunday)
        Case vbSaturday
            CollectionDate = CDate(Date - Weekday(Date, vbSunday) + 9)
            'Debug.Print Weekday(Date, vbSunday)
        Case vbSunday
            CollectionDate = CDate(Date - Weekday(Date, vbSunday) + 2)
            'Debug.Print Weekday(Date, vbSunday)
        Case Else
            CollectionDate = CDate(Date)
        
    End Select
End Function
Private Sub GenerateOrderNo()
    Dim TodayDate As Date
    
    TodayDate = Format(Date, "dd/mm/yyyy")
    
    Select Case Weekday(Date)
            Case vbMonday, vbWednesday, vbFriday, vbSaturday
                If wsSettings.Range("Q2").Value <> CDate(TodayDate) Then
                    Me.OrderNo.Value = wsSettings.Range("P2").Value + 1
                    wsSettings.Range("P2").Value = Me.OrderNo.Value
                    wsSettings.Range("Q2").Value = Date
                End If
            Case Else
    End Select
    
End Sub
'Private Sub GenerateRandomNumber()
'    Dim min As Long
'    Dim max As Long
'    Dim randomUniqueNumberList As Long
'    Dim i As Long
'    Dim uniqueNumber As Long
'    Dim orderNoRange As Range
'    'Dim dateFormat As String
'
'
'    min = 10000
'    max = 99999
'    'dateFormat = Format(Now, "dd/mm/yyyy")
'    randomUniqueNumberList = wsSettings.Range("R" & wsSettings.Rows.Count).End(xlUp).Row
'
'    'Change the below when creating the sheet names for staff to qualify the range for the specific sheet
'    Set orderNoRange = wsSettings.Range("S2:S" & randomUniqueNumberList)
'
'    For i = 2 To randomUniqueNumberList
'            If Not wsSettings.Range("S2").value = Date Then
'
'                If OrderNo.value = 0 Or wsSettings.Range("D2").value = 0 Then
'
'                If uniqueNumber = 0 Then uniqueNumber = Int((max - min + 1) * Rnd + 1)
'
'                    If IsUnique(uniqueNumber, orderNoRange) Then
'                        OrderNo.value = uniqueNumber
'                        wsSettings.Range("D2").value = uniqueNumber
'                        wsSettings.Range("E2").value = Date
'                        Exit Sub
'                    Else
'                        uniqueNumber = Int((max - min + 1) * Rnd + 1)
'                    End If
'                End If
'            Else
'                Exit Sub
'            End If
'    Next i
'End Sub
'Function IsUnique(Num As Long, RangeList As Variant) As Boolean
'    Dim iFind As Long
'    'IsUnique = True
'    'part of random no
'
'    On Error GoTo Unique
'    iFind = Application.WorksheetFunction.Match(Num, RangeList, 0)
'
'    If iFind > 0 Then IsUnique = False: Exit Function
'
'Unique:
'    IsUnique = True
'End Function

Private Sub UserForm_Initialize()
    
    'Hide workbook when launched
    SmcDryCleaningForm = ThisWorkbook.Name
    wsAdmin.Visible = xlSheetVeryHidden
    
    If Workbooks.Count > 1 Then
            Windows(SmcDryCleaningForm).Visible = False
        Else
            Application.Visible = False
        End If
     
    'Setting Dynamic name range in the drop down list
    EmployeeDropDownList.List = wsSettings.Range("A2", wsSettings.Range("A2").End(xlDown)).Value
    NewEmployeePosition.List = wsSettings.Range("N2", wsSettings.Range("N2").End(xlDown)).Value
    'Call RemoveItemFromList
    
    'Call FillInItemsToClean '<--- no need to load uniform until employee selection
    Me.OrderNo.Value = wsSettings.Range("P2").Value
    Me.CurrentDate.Value = WeekdayName(Weekday(Now - 1), False) & " " & Format(Date, "dd/mm/yyyy")
    Me.NextCollectionDate.Value = Format(CollectionDate(Date), "dd/mm/yyyy")
    Me.DateReturnBy.Value = Format(CollectionDate(Date), "dd/mm/yyyy")
End Sub
Private Sub FillInItemsToClean()

    Dim i As Long
    Dim cmb As ComboBox
    Dim UniformOption As String
    Dim UniformSelection As Variant
    
    If PositionTitle.Value <> vbNullString Then
    
    UniformOption = PositionTitle.Value
    
        Select Case UniformOption
               Case "Store Manager", "SM"
                    UniformSelection = wsSettings.Range("D2", wsSettings.Range("D2").End(xlDown)).Value
               Case "Assistant Store Manager", "ASM"
                    UniformSelection = wsSettings.Range("F2", wsSettings.Range("F2").End(xlDown)).Value
               Case "Stock Controller", "SC"
                    UniformSelection = wsSettings.Range("H2", wsSettings.Range("H2").End(xlDown)).Value
               Case "Sales Assistant", "SA"
                    UniformSelection = wsSettings.Range("J2", wsSettings.Range("J2").End(xlDown)).Value
               Case "Other"
                    UniformSelection = wsSettings.Range("L2", wsSettings.Range("L2").End(xlDown)).Value
               Case Else
                    
        End Select
        
        For i = 1 To 13
            Set cmb = Me.Controls("Item" & i)
            cmb.List = UniformSelection
        Next i
    End If

End Sub
Private Sub EmployeeDropDownList_DropButtonClick()
Dim FoundEmployee As Range

    EmployeeDropDownList.List = wsSettings.Range("A2", wsSettings.Range("A2").End(xlDown)).Value
    With wsSettings.Range("A:A")
        Set FoundEmployee = .Cells.Find(What:=EmployeeDropDownList.Value, after:=.Cells(1, 1), _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    End With
    
    PositionTitle.Value = FoundEmployee.Offset(0, 1)
    
    Call FillInItemsToClean
    
    'OrderNo.Value = wsSettings.Range("S2").Value 'set a new order no incrementingby one each collection
End Sub
Private Sub UpdateEmployeeItemsToClean()
    Dim i As Long
    Dim WS As Worksheet
    
    Dim ctl As MSForms.Control
    Dim lbl As MSForms.Label
    Dim cmb As MSForms.ComboBox
    Dim txtbox As MSForms.TextBox
    Dim optbtn As MSForms.OptionButton
    
    Set WS = ActiveSheet
    
With WS
    'For i = 1 To ItemsListFrame.Controls.Count ' 125 controls so no point to loop through the whole count
    For i = 1 To 13
        For Each ctl In ItemsListFrame.Controls
            If TypeName(ctl) = "Label" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set lbl = ctl
                    If .Range("A" & i + 6).Offset(0, 2).Value <> vbNullString Then
                        If Controls("Item" & i).Value = vbNullString Then
                            Controls("OrderLbl" & i).Enabled = True
                            '.Range("A" & i + 6).Offset(0, 0).Value = Me.OrderNo.Value
                            '.Range("A" & i + 6).Offset(0, 1).Value = Me.NextCollectionDate.Text
                            '.Range("A" & i + 6).Offset(0, 1).Value = Format(.Range("A" & i + 6).Offset(0, 1).Value, "dd/mm/yyyy")
                            '.Range("A" & i + 6).Offset(0, 8).Value = Me.DateReturnBy.Value
                            '.Range("A" & i + 6).Offset(0, 8).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                        End If
                    Else
                        Controls("OrderLbl" & i).Enabled = False
                    End If
                End If
    
            ElseIf TypeName(ctl) = "ComboBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set cmb = ctl
                    If .Range("A" & i + 6).Offset(0, 2).Value <> vbNullString Then
                        If Controls("Item" & i).Value = vbNullString Or Controls("Item" & i).Value <> vbNullString Then
                            Controls("Item" & i).Enabled = True
                            Controls("Item" & i).Value = .Range("A" & i + 6).Offset(0, 2).Value
                            End If
                    Else
                        Controls("Item" & i).Value = ""
                        Controls("Item" & i).Enabled = False
                    End If
                End If
    
            ElseIf TypeName(ctl) = "TextBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set txtbox = ctl
                    If .Range("A" & i + 6).Offset(0, 3).Value <> vbNullString Or .Range("A" & i + 6).Offset(0, 4).Value <> vbNullString Or .Range("A" & i + 6).Offset(0, 5).Value <> vbNullString Or .Range("A" & i + 6).Offset(0, 7).Value <> vbNullString Or .Range("A" & i + 6).Offset(0, 8).Value <> vbNullString Then
                        If Controls("Item" & i).Value <> vbNullString Then
                            Controls("Qty" & i).Enabled = True
                            Controls("UnitPrice" & i).Enabled = True
                            Controls("SubTotal" & i).Enabled = True
                            Controls("Comments" & i).Enabled = True
                            Controls("ReturnDate" & i).Enabled = True
                            Controls("Qty" & i).Value = .Range("A" & i + 6).Offset(0, 3).Value
                            Controls("UnitPrice" & i).Value = .Range("A" & i + 6).Offset(0, 4).Value
                            Controls("SubTotal" & i).Value = .Range("A" & i + 6).Offset(0, 5).Value
                            Controls("Comments" & i).Value = .Range("A" & i + 6).Offset(0, 7).Value ' & " // Return by " & .Range("A" & i + 6).Offset(0, 8).Value
                            Controls("ReturnDate" & i).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                            
                            If .Range("A" & i + 6).Offset(0, 6).Value = "OUT" Then
                                .Range("A" & i + 6).Offset(0, 8).Value = Format(Controls("ReturnDate" & i).Value, "dd/mm/yyyy")
                                Controls("ReturnDate" & i).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                            Else
                                .Range("A" & i + 6).Offset(0, 8).Value = ""
                                Controls("ReturnDate" & i).Value = ""
                            End If
                            
                            
                        End If
                    Else
                        Controls("Qty" & i).Enabled = False
                        Controls("UnitPrice" & i).Enabled = False
                        Controls("SubTotal" & i).Enabled = False
                        Controls("Comments" & i).Enabled = False
                        Controls("ReturnDate" & i).Enabled = False
                        Controls("Qty" & i).Value = 1
                        If .Range("A" & i + 6).Offset(0, 4).Value = vbNullString Then
                            Controls("UnitPrice" & i).Value = "0.00"
                        Else
                            Controls("UnitPrice" & i).Value = .Range("A" & i + 6).Offset(0, 4).Value
                        End If
                        If .Range("A" & i + 6).Offset(0, 5).Value = vbNullString Then
                            Controls("SubTotal" & i).Value = "0.00"
                        Else
                            Controls("SubTotal" & i).Value = .Range("A" & i + 6).Offset(0, 5).Value
                        End If
                        Controls("Comments" & i).Value = ""
                        If .Range("A" & i + 6).Offset(0, 6).Value = "IN" Or .Range("A" & i + 6).Offset(0, 6).Value = vbNullString Then
                            Controls("ReturnDate" & i).Value = ""
                            .Range("A" & i + 6).Offset(0, 8).Value = ""
                        'Else
                            'Controls("ReturnDate" & i).Value = .Range("A" & i + 6).Offset(0, 8).Value
                        End If
                    End If
                End If
            ElseIf TypeName(ctl) = "OptionButton" Then
                If ctl.Tag = "GroupItem" & i Or ctl.Tag = "InOut" & i Then
                    Set optbtn = ctl
                    If .Range("A" & i + 6).Offset(0, 6).Value <> vbNullString Then
                        If Controls("Item" & i).Value <> vbNullString Then
                        
                            If .Range("A" & i + 6).Offset(0, 6).Value = "OUT" Then
                                Controls("OptionIn" & i).Enabled = True
                                Controls("OptionOut" & i).Enabled = True
                                Controls("OptionOut" & i).Value = True
                            Else
                                Controls("OptionIn" & i).Enabled = True
                                Controls("OptionOut" & i).Enabled = True
                                Controls("OptionIn" & i).Value = True
                            End If
                        End If
                    Else
                        Controls("OptionIn" & i).Enabled = False
                        Controls("OptionOut" & i).Enabled = False
                        Controls("OptionIn" & i).Value = True
                    End If
                End If
            End If
        Next ctl
      Next i
End With
End Sub
Private Sub AddNewEmployeeBtnPressed_Click()

    Me.NewEmployeeFrame.Visible = True
    Me.NewEmployeeFirstName.SetFocus
    
    FormDataMode = AddNewEmployeeBtnPressed

End Sub
Private Sub NewEmployeeFirstName_AfterUpdate()
    
    If NewEmployeeFirstName.Value <> "" Then
        NewEmployeeFirstNameLbl.ForeColor = Me.ForeColor
        NewEmployeeFirstName.BackColor = rgbWhite
        'If EverythingFilledIn = True Then NewEmployeeMessageErrorLbl.Caption = ""
    End If
End Sub
Private Sub NewEmployeeLastName_AfterUpdate()

    If NewEmployeeLastName.Value <> "" Then
        NewEmployeeLastNameLbl.ForeColor = Me.ForeColor
        NewEmployeeLastName.BackColor = rgbWhite
        NewEmployeeLastName.SetFocus
        'If EverythingFilledIn = True Then NewEmployeeMessageErrorLbl.Caption = ""
    End If
End Sub
Private Sub NewEmployeePosition_AfterUpdate()

    If NewEmployeePosition.Value <> "" Then
        NewEmployeePositionLbl.ForeColor = Me.ForeColor
        NewEmployeePosition.BackColor = rgbWhite
        NewEmployeePosition.SetFocus
        If EverythingFilledIn = True Then NewEmployeeMessageErrorLbl.Caption = ""
    End If
End Sub
Private Sub AddNewEmployeeToListBtnPressed_Click()
    
    If Not EverythingFilledIn Then Exit Sub
    
    If Not UniqueEmployeeNames Then
        Me.NewEmployeeMessageErrorLbl.Caption = "The employee name entered exists. Please change it and try again."
        NewEmployeeFirstNameLbl.ForeColor = rgbRed
        NewEmployeeFirstName.BackColor = rgbPink
        NewEmployeeLastNameLbl.ForeColor = rgbRed
        NewEmployeeLastName.BackColor = rgbPink
        NewEmployeePositionLbl.ForeColor = rgbRed
        NewEmployeePosition.BackColor = rgbPink
        Exit Sub
    End If
    
    Call AddEmployeeToList
       
    Me.NewEmployeeMessageErrorLbl.Caption = ""
    Me.NewEmployeeFrame.Visible = False

    FormDataMode = AllBtnsEnabled
    
End Sub
Private Function UniqueEmployeeNames() As Boolean
    
    Dim i As Long
    Dim employeeList As Long

    UniqueEmployeeNames = True
    employeeList = wsSettings.Range("A" & wsSettings.Rows.Count).End(xlUp).Row
    
    For i = 2 To employeeList
        If wsSettings.Range("A" & i).Value = Application.WorksheetFunction.Proper(NewEmployeeFirstName.Value & " " & NewEmployeeLastName.Value) Then
            UniqueEmployeeNames = False
        End If
    Next i
    
End Function
Private Function EverythingFilledIn() As Boolean

    Dim ctl As MSForms.Control
    Dim txt As MSForms.TextBox
    Dim cmb As MSForms.ComboBox
    Dim AnythingIsMissing As Boolean
    
    EverythingFilledIn = True
    AnythingIsMissing = False
    
    For Each ctl In NewEmployeeFrame.Controls
        If TypeOf ctl Is MSForms.TextBox Then
            Set txt = ctl
            If txt.Value = "" Then
                txt.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.NewEmployeeMessageErrorLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledIn = False
                    If EverythingFilledIn = True Then NewEmployeeMessageErrorLbl.Caption = ""
            End If
        End If
        
            If TypeOf ctl Is MSForms.ComboBox Then
            Set cmb = ctl
            If cmb.Value = "" Then
                cmb.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.NewEmployeeMessageErrorLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledIn = False
                    If EverythingFilledIn = True Then NewEmployeeMessageErrorLbl.Caption = ""
            End If
        End If
    Next ctl
    
Set ctl = Nothing

End Function
Private Sub AddEmployeeToList()
    Dim lrow As Long
    
    lrow = wsSettings.Range("A" & wsSettings.Rows.Count).End(xlUp).Row + 1
    wsSettings.Range("A" & lrow).Value = Application.WorksheetFunction.Proper(NewEmployeeFirstName.Value & " " & NewEmployeeLastName.Value)
    wsSettings.Range("B" & lrow).Value = Application.WorksheetFunction.Proper(NewEmployeePosition.Value)
    
    'to check if workbook name exist
    If Not sheetExists(Me.NewEmployeeFirstName.Value & " " & Me.NewEmployeeLastName.Value) Then
        Call CreateNewEmployeeSheet
    End If
    
    
    NewEmployeeFirstName.Value = ""
    NewEmployeeLastName.Value = ""
    NewEmployeePosition.Value = ""
        
End Sub
Private Sub CreateNewEmployeeSheet()
    
    Dim WS As Worksheet
    
    wsTemplate.Copy after:=Sheets(3)
    Set WS = ActiveSheet
    WS.Name = Me.NewEmployeeFirstName.Value & " " & Me.NewEmployeeLastName.Value
    WS.Range("B1").Value = Me.NewEmployeeFirstName.Value
    WS.Range("B2").Value = Me.NewEmployeeLastName.Value
    WS.Range("B3").Value = Me.NewEmployeePosition.Value
    WS.Range("B4").Value = Me.CurrentDate.Value
    
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
Sub ActivateSheet(SheetName As String)

    Worksheets(SheetName).Activate
End Sub

Private Sub NewEmployeeCancelBtn_Click()

    NewEmployeeFirstName.Value = ""
    NewEmployeeLastName.Value = ""
    NewEmployeePosition.Value = ""
    NewEmployeeFirstNameLbl.ForeColor = Me.ForeColor
    NewEmployeeFirstName.BackColor = rgbWhite
    NewEmployeeLastNameLbl.ForeColor = Me.ForeColor
    NewEmployeeLastName.BackColor = rgbWhite
    NewEmployeePositionLbl.ForeColor = Me.ForeColor
    NewEmployeePosition.BackColor = rgbWhite
    NewEmployeeMessageErrorLbl.Caption = ""
    Me.NewEmployeeFrame.Visible = False

    FormDataMode = AllBtnsEnabled
    
End Sub
Private Sub RemoveEmployeeBtnPressed_Click()

        Me.RemoveEmployeeFrame.Visible = True
        Me.RemoveEmployeeDropDownList.SetFocus
        
        FormDataMode = RemoveEmployeeBtn
End Sub
Private Sub RemoveEmployeeDropDownList_DropButtonClick()
Dim FoundEmployee As Range

    RemoveEmployeeDropDownList.List = wsSettings.Range("A2", wsSettings.Range("A2").End(xlDown)).Value
     With wsSettings.Range("A:A")
        Set FoundEmployee = .Cells.Find(What:=RemoveEmployeeDropDownList.Value, after:=.Cells(1, 1), _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    End With
    RemoveEmployeePosition.Value = FoundEmployee.Offset(0, 1)
End Sub
Private Sub RemoveEmployeeDropDownList_AfterUpdate()

    If RemoveEmployeeDropDownList.Value <> "" Then
        Me.RemoveEmployeeDropDownListLbl.ForeColor = Me.ForeColor
        Me.RemoveEmployeeDropDownList.BackColor = vbWhite
        Me.RemoveEmployeePositionLbl.ForeColor = Me.ForeColor
        Me.RemoveEmployeePosition.BackColor = vbWhite
        Me.RemoveEmployeePosition.SetFocus
        
        If EverythingFilledInRemoveEmployee = True Then RemoveEmployeeErrorMsgLbl.Caption = ""
    End If
    
End Sub
Private Sub RemoveEmployeeRecordBtnPressed_Click()

    Call RemoveEmployeeFromList
    
    If Not EverythingFilledInRemoveEmployee Then Exit Sub
    
    If Me.RemoveEmployeeDropDownList.Value = vbNullString Then Exit Sub
    
    Me.RemoveEmployeeDropDownList.Value = ""
    Me.RemoveEmployeePosition.Value = ""
    Me.RemoveEmployeeFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
End Sub
Private Function EverythingFilledInRemoveEmployee() As Boolean

    Dim ctl As MSForms.Control
    Dim txt As MSForms.TextBox
    Dim cmb As MSForms.ComboBox
    Dim AnythingIsMissing As Boolean
    
    EverythingFilledInRemoveEmployee = True
    AnythingIsMissing = False
    
    For Each ctl In RemoveEmployeeFrame.Controls
        If TypeOf ctl Is MSForms.TextBox Then
            Set txt = ctl
            If txt.Value = "" Then
                txt.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.RemoveEmployeeErrorMsgLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledInRemoveEmployee = False
                    If EverythingFilledInRemoveEmployee = True Then RemoveEmployeeErrorMsgLbl.Caption = ""
            End If
        End If
        
            If TypeOf ctl Is MSForms.ComboBox Then
            Set cmb = ctl
            If cmb.Value = "" Then
                cmb.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.RemoveEmployeeErrorMsgLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledInRemoveEmployee = False
                    If EverythingFilledInRemoveEmployee = True Then RemoveEmployeeErrorMsgLbl.Caption = ""
            End If
        End If
    Next ctl
    
Set ctl = Nothing

End Function
Private Sub RemoveEmployeeFromList()

    Dim lrow As Long, i As Long
    Dim WS As Worksheet
    
    Set WS = wsSettings
    With WS
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row
        
        For i = 2 To lrow
            
            'To use CStr when comparing values in cells not any other method as you can't guarante what type of value is in the cell
            If CStr(Len(Trim(.Range("A" & i).Text))) = CStr(Len(Trim(RemoveEmployeeDropDownList.Text))) Then
                .Range("A" & i, "B" & i).Delete
            End If
        Next i

    End With
End Sub
Private Sub RemoveEmployeeCancelBtn_Click()

    Me.RemoveEmployeeDropDownList.Value = ""
    Me.RemoveEmployeePosition.Value = ""
    Me.RemoveEmployeeDropDownListLbl.ForeColor = Me.ForeColor
    Me.RemoveEmployeeDropDownList.BackColor = vbWhite
    Me.RemoveEmployeePositionLbl.ForeColor = Me.ForeColor
    Me.RemoveEmployeePosition.BackColor = vbWhite

    Me.RemoveEmployeeFrame.Visible = False
    Me.RemoveEmployeeErrorMsgLbl.Caption = ""
    
    FormDataMode = AllBtnsEnabled
End Sub
Private Sub AddNewItemBtnPressed_Click()

    Me.NewItemFrame.Visible = True
    Me.AddNewItemPosition.SetFocus
    
    FormDataMode = AddNewItemBtn
End Sub
Private Sub AddNewItemPosition_DropButtonClick()
        
        Call UniquePositionTitle
End Sub
Private Sub UniquePositionTitle()

    Dim rng As Range
    Dim Dn As Range, arr, v

    Set rng = wsSettings.Range("B2", wsSettings.Cells(Rows.Count, "B").End(xlUp))

    Set Dic = CreateObject("scripting.dictionary")
    Dic.CompareMode = vbTextCompare
    For Each Dn In rng
        v = Dn.Offset(0, 1)

        If Not Dic.exists(Dn.Value) Then
            Dic.Add Dn.Value, Array(v)
        Else
            arr = Dic(Dn.Value)
            'no match will return an error value: test for this
            If IsError(Application.Match(v, arr, 0)) Then
                ReDim Preserve arr(UBound(arr) + 1)
                arr(UBound(arr)) = v
                Dic(Dn.Value) = arr 'replace with expanded array
            End If
        End If

    Next

    AddNewItemPosition.List = Dic.Keys
    RemoveItemPosition.List = Dic.Keys
    
    Set Dic = Nothing
End Sub
Private Sub NewItemBtnPressed_Click()

    If Not EverythingFilledInItem Then Exit Sub
    
    If Not UniqueItemNames Then
        Me.AddNewItemMsgLbl.Caption = "The item entered exists. Please change it and try again."
        AddNewItemPositionLbl.ForeColor = Me.ForeColor
        AddNewItemPosition.BackColor = rgbWhite
        AddNewItemLbl.ForeColor = rgbRed
        AddNewItem.BackColor = rgbPink
        Exit Sub
    End If
    
    Call AddNewItemsToList
    
    AddNewItemPosition.Value = ""
    AddNewItem.Value = ""
    
    Me.AddNewItemMsgLbl.Caption = ""
    Me.NewItemFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
    
End Sub
Private Function EverythingFilledInItem() As Boolean

    Dim ctl As MSForms.Control
    Dim txt As MSForms.TextBox
    Dim cmb As MSForms.ComboBox
    Dim AnythingIsMissing As Boolean
    
    EverythingFilledInItem = True
    AnythingIsMissing = False
    
    For Each ctl In NewItemFrame.Controls
        If TypeOf ctl Is MSForms.ComboBox Then
            Set cmb = ctl
            If cmb.Value = "" Then
                cmb.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.AddNewItemMsgLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledInItem = False
                    If EverythingFilledInItem = True Then AddNewItemMsgLbl.Caption = ""
            End If
        End If
        
        If TypeOf ctl Is MSForms.TextBox Then
            Set txt = ctl
            If txt.Value = "" Then
                txt.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.AddNewItemMsgLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then ctl.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledInItem = False
                    If EverythingFilledInItem = True Then AddNewItemMsgLbl.Caption = ""
            End If
        End If
    Next ctl
    
Set ctl = Nothing

End Function
Private Sub AddNewItemPosition_AfterUpdate()

    If AddNewItemPosition.Value <> "" Then
        AddNewItemPositionLbl.ForeColor = Me.ForeColor
        AddNewItemPosition.BackColor = rgbWhite
'        AddNewItemLbl.ForeColor = Me.ForeColor
'        AddNewItem.BackColor = rgbWhite
        AddNewItem.SetFocus
        'If EverythingFilledInItem = True Then AddNewItemMsgLbl.Caption = ""
    End If
End Sub
Private Sub AddNewItem_AfterUpdate()

    If AddNewItem.Value <> "" Then
        AddNewItemLbl.ForeColor = Me.ForeColor
        AddNewItem.BackColor = rgbWhite
        AddNewItem.SetFocus
        If EverythingFilledInItem = True Then AddNewItemMsgLbl.Caption = ""
    End If
End Sub
Private Sub AddNewItemsToList()
    Dim lrow As Long
    
    If AddNewItemPosition <> "" Then
        If AddNewItemPosition.Value = "Store Manager" Or AddNewItemPosition.Value = "SM" Then
            lrow = wsSettings.Range("D" & wsSettings.Rows.Count).End(xlUp).Row + 1
            Range("D" & lrow).Value = Application.WorksheetFunction.Proper(AddNewItem.Value)
        ElseIf AddNewItemPosition.Value = "Assistant Store Manager" Or AddNewItemPosition.Value = "ASM" Then
            lrow = wsSettings.Range("F" & wsSettings.Rows.Count).End(xlUp).Row + 1
            Range("F" & lrow).Value = Application.WorksheetFunction.Proper(AddNewItem.Value)
        ElseIf AddNewItemPosition.Value = "Stock Controller" Or AddNewItemPosition.Value = "SC" Then
            lrow = wsSettings.Range("H" & wsSettings.Rows.Count).End(xlUp).Row + 1
            Range("H" & lrow).Value = Application.WorksheetFunction.Proper(AddNewItem.Value)
        ElseIf AddNewItemPosition.Value = "Sales Assistant" Or AddNewItemPosition.Value = "SA" Then
            lrow = wsSettings.Range("J" & wsSettings.Rows.Count).End(xlUp).Row + 1
            Range("J" & lrow).Value = Application.WorksheetFunction.Proper(AddNewItem.Value)
        Else
         
            lrow = wsSettings.Range("L" & wsSettings.Rows.Count).End(xlUp).Row + 1
            Range("L" & lrow).Value = Application.WorksheetFunction.Proper(AddNewItem.Value)
        End If
    End If
End Sub
Private Function UniqueItemNames() As Boolean
    
    Dim i As Long
    Dim ItemList As Long
    Dim nCol As String

    UniqueItemNames = True
    
    If AddNewItemPosition.Value = "Store Manager" Or AddNewItemPosition.Value = "SM" Then
            ItemList = wsSettings.Range("D" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "D"
        ElseIf AddNewItemPosition.Value = "Assistant Store Manager" Or AddNewItemPosition.Value = "ASM" Then
            ItemList = wsSettings.Range("F" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "F"
        ElseIf AddNewItemPosition.Value = "Stock Controller" Or AddNewItemPosition.Value = "SC" Then
            ItemList = wsSettings.Range("H" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "H"
        ElseIf AddNewItemPosition.Value = "Sales Assistant" Or AddNewItemPosition.Value = "SA" Then
            ItemList = wsSettings.Range("J" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "J"
        Else
            ItemList = wsSettings.Range("L" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "L"
        End If
    
    For i = 2 To ItemList
        If wsSettings.Range(nCol & i).Value = Application.WorksheetFunction.Proper(AddNewItem.Value) Then
            UniqueItemNames = False
        End If
    Next i
    
End Function
Private Sub NewItemCancelBtnPressed_Click()

    Me.AddNewItemPosition.Value = ""
    Me.AddNewItem.Value = ""
    Me.AddNewItemPositionLbl.ForeColor = Me.ForeColor
    Me.AddNewItemPosition.BackColor = rgbWhite
    Me.AddNewItemLbl.ForeColor = Me.ForeColor
    Me.AddNewItem.BackColor = rgbWhite
    Me.AddNewItemMsgLbl = ""
    Me.NewItemFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
    
End Sub
Private Sub RemoveItemBtnPress_Click()
    
    Me.RemoveItemFrame.Visible = True
    Me.RemoveItemPosition.SetFocus
    
    FormDataMode = RemoveItemBtn
End Sub
Private Sub RemoveItemPosition_DropButtonClick()

    Call UniquePositionTitle
End Sub
Private Sub RemoveItem_DropButtonClick()

    Call SelectItemFromDropDownList
End Sub
Private Sub SelectItemFromDropDownList()

    Dim lrow As Long
    Dim WS As Worksheet
    Dim nCol As String
    
    Set WS = wsSettings
    With WS
    
    If RemoveItemPosition.Value = "Store Manager" Or RemoveItemPosition.Value = "SM" Then
            lrow = wsSettings.Range("D" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "D"
        ElseIf RemoveItemPosition.Value = "Assistant Store Manager" Or RemoveItemPosition.Value = "ASM" Then
            lrow = wsSettings.Range("F" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "F"
        ElseIf RemoveItemPosition.Value = "Stock Controller" Or RemoveItemPosition.Value = "SC" Then
            lrow = wsSettings.Range("H" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "H"
        ElseIf RemoveItemPosition.Value = "Sales Assistant" Or RemoveItemPosition.Value = "SA" Then
            lrow = wsSettings.Range("J" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "J"
        Else
            lrow = wsSettings.Range("L" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "L"
        End If
        
        RemoveItem.List = wsSettings.Range(nCol & 2, wsSettings.Cells(Rows.Count, nCol).End(xlUp)).Value

    End With
End Sub
Private Sub RemoveItemFromList()

    Dim lrow As Long, i As Long
    Dim WS As Worksheet
    Dim nCol As String
    
    Set WS = wsSettings
    With WS
    
    If RemoveItemPosition.Value = "Store Manager" Or RemoveItemPosition.Value = "SM" Then
            lrow = wsSettings.Range("D" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "D"
        ElseIf RemoveItemPosition.Value = "Assistant Store Manager" Or RemoveItemPosition.Value = "ASM" Then
            lrow = wsSettings.Range("F" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "F"
        ElseIf RemoveItemPosition.Value = "Stock Controller" Or RemoveItemPosition.Value = "SC" Then
            lrow = wsSettings.Range("H" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "H"
        ElseIf RemoveItemPosition.Value = "Sales Assistant" Or RemoveItemPosition.Value = "SA" Then
            lrow = wsSettings.Range("J" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "J"
        Else
            lrow = wsSettings.Range("L" & wsSettings.Rows.Count).End(xlUp).Row
            nCol = "L"
        End If
        
        RemoveItem.List = wsSettings.Range(nCol & 2, wsSettings.Cells(Rows.Count, nCol).End(xlUp)).Value
        
        For i = 2 To lrow
            
            'To use CStr when comparing values in cells not any other method as you can't guarante what type of value is in the cell
            'If StrComp(CStr(UCase(Len(Trim(.Range(nCol & i).Text)))), CStr(UCase(Len(Trim(RemoveItem.Text)))), vbTextCompare) = 0 Then
            If StrComp(CStr(UCase(Trim(.Range(nCol & i).Text))), CStr(UCase(Trim(RemoveItem.Text))), vbTextCompare) = 0 Then
                .Range(nCol & i).Delete
            End If
        Next i

    End With
End Sub
Private Sub RemoveItemBtnPressed_Click()

    If Not EverythingFilledInRemoveItem Then Exit Sub
    
    If Not UniqueItemNames Then
        Me.RemoveItemMsgLbl.Caption = "The item entered exists. Please change it and try again."
        RemoveItemPositionLbl.ForeColor = Me.ForeColor
        RemoveItemPosition.BackColor = rgbWhite
        RemoveItemLbl.ForeColor = rgbRed
        RemoveItem.BackColor = rgbPink
        Exit Sub
    End If
    
    Call RemoveItemFromList
    
    RemoveItemPosition.Value = ""
    RemoveItem.Value = ""
    
    Me.RemoveItemMsgLbl.Caption = ""
    Me.RemoveItemFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
    
End Sub
Private Function EverythingFilledInRemoveItem() As Boolean

    Dim ctl As MSForms.Control
    Dim txt As MSForms.TextBox
    Dim cmb As MSForms.ComboBox
    Dim AnythingIsMissing As Boolean
    
    EverythingFilledInRemoveItem = True
    AnythingIsMissing = False
    
    For Each ctl In RemoveItemFrame.Controls
        If TypeOf ctl Is MSForms.ComboBox Then
            Set cmb = ctl
            If cmb.Value = "" Then
                cmb.BackColor = rgbPink
                Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                Me.RemoveItemMsgLbl.Caption = "One or more fields are empty. Please try again."
                If Not AnythingIsMissing Then cmb.SetFocus
                    AnythingIsMissing = True
                    EverythingFilledInRemoveItem = False
                    If EverythingFilledInRemoveItem = True Then RemoveItemMsgLbl.Caption = ""
            End If
        End If
    Next ctl
    
Set ctl = Nothing

End Function
Private Sub RemoveItemPosition_AfterUpdate()

    If RemoveItemPosition.Value <> "" Then
        RemoveItemPositionLbl.ForeColor = Me.ForeColor
        RemoveItemPosition.BackColor = rgbWhite
        RemoveItem.SetFocus
        'If EverythingFilledInRemoveItem = True Then AddNewItemMsgLbl.Caption = ""
    End If
End Sub
Private Sub RemoveItem_AfterUpdate()

    If RemoveItem.Value <> "" Then
        RemoveItemLbl.ForeColor = Me.ForeColor
        RemoveItem.BackColor = rgbWhite
        RemoveItem.SetFocus
        If EverythingFilledInRemoveItem = True Then RemoveItemMsgLbl.Caption = ""
    End If
End Sub
Private Sub RemoveItemCancelBtnPressed_Click()
    
    Me.RemoveItemPosition.Value = ""
    Me.RemoveItem.Value = ""
    Me.RemoveItemPositionLbl.ForeColor = Me.ForeColor
    Me.RemoveItemPosition.BackColor = rgbWhite
    Me.RemoveItemLbl.ForeColor = Me.ForeColor
    Me.RemoveItem.BackColor = rgbWhite
    Me.RemoveItemMsgLbl = ""
    Me.RemoveItemFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
    
End Sub

Private Sub ShowHideTogglePressed_Click()

  With Me.ShowHideTogglePressed
    If .Value = True Then
        .Caption = "Hide Records"
        If Workbooks.Count > 1 Then
            Windows(SmcDryCleaningForm).Visible = True
        Else
            Application.Visible = True
        End If
    Else
        .Caption = "Show Records"
        If Workbooks.Count > 1 Then
            Windows(SmcDryCleaningForm).Visible = False
        Else
            Application.Visible = False
        End If
    End If
  End With
End Sub
Private Sub AdminModeBtnPressed_Click()

    Me.AdminModeFrame.Visible = True
    Me.AdminModeLoginAddBtn.Caption = "Login"
    Me.AdminModeAdminName.SetFocus
    
    Me.AdminModeNewUserProfileLbl.Visible = False
    Me.AdminModeNewUserProfile.Visible = False
    
    FormDataMode = AdminModeBtn
End Sub
Private Sub AdminModeLoginAddBtn_Click()

    If Not EverythingFilledInAdminMode Then Exit Sub '<--maybe not needing this
    
    If Me.AdminModeLoginAddBtn.Caption = "Login" Then
        Call LoginUser

    Else
        Call AddUserProfile
        
        If isUniqueUser(Me.AdminModeNewUserProfile) = True Then
            Me.AdminModeAdminName.Value = ""
            Me.AdminModePassword.Value = ""
            Me.AdminModeNewUserProfile.Value = ""
            
            Me.AdminModeMessageLbl.Caption = ""
            Me.AdminModeFrame.Visible = False
            
            FormDataMode = AllBtnsEnabled
        Else
            Me.AdminModeMessageLbl.Caption = "User already exist. Please try again!"
            AdminModeNewUserProfileLbl.ForeColor = rgbRed
            AdminModeNewUserProfile.BackColor = rgbPink
            AdminModeNewUserProfile.SetFocus
        End If
        
    End If

End Sub
Private Sub AdminModeAdminName_AfterUpdate()

    If AdminModeAdminName.Value <> "" Then
        AdminModeAdminNameLbl.ForeColor = Me.ForeColor
        AdminModeAdminName.BackColor = rgbWhite
'        AddNewItemLbl.ForeColor = Me.ForeColor
'        AddNewItem.BackColor = rgbWhite
        AdminModeAdminName.SetFocus
        'If EverythingFilledInAdminMode = True Then AddNewItemMsgLbl.Caption = ""
    End If
End Sub
Private Sub AdminModePassword_AfterUpdate()

    If AdminModePassword.Value <> "" Then
        AdminModePasswordLbl.ForeColor = Me.ForeColor
        AdminModePassword.BackColor = rgbWhite
        AdminModePassword.SetFocus
        If EverythingFilledInAdminMode = True Then AdminModeMessageLbl.Caption = ""
    End If
End Sub
Private Sub AdminModeNewUserProfile_AfterUpdate()
    
    If isUniqueUser(Me.AdminModeNewUserProfile) = True Then
        AdminModeNewUserProfileLbl.ForeColor = Me.ForeColor
        AdminModeNewUserProfile.BackColor = rgbWhite
        AdminModeNewUserProfile.SetFocus
        If EverythingFilledInAdminMode = True Then AdminModeMessageLbl.Caption = ""
    End If
End Sub
Private Sub LoginUser()
    Dim Username As String
    Dim password As String
    Dim passWs As Worksheet
    Dim rng As Range
    Dim lrow As Long
    Dim i As Long
    
    Username = AdminModeAdminName.Text
    password = AdminModePassword.Text
    CorrectLoginDetails = False
    
    If Len(Trim(Username)) = 0 Then
        AdminModeAdminName.SetFocus
        Me.AdminModeMessageLbl.Caption = "Please enter a username!"
        Exit Sub
    End If

    If Len(Trim(password)) = 0 Then
        AdminModePassword.SetFocus
        Me.AdminModeMessageLbl.Caption = "Please enter a password!"
        Exit Sub
    End If

    Set passWs = ThisWorkbook.Worksheets("Admin")

    With passWs
        lrow = .Range("A" & .Rows.Count).End(xlUp).Row

        For i = 2 To lrow
            If UCase(Trim(.Range("A" & i).Value)) = UCase(Trim(Username)) Then '<~~ Username Check
                If .Range("B" & i).Value = password Then '<~~ Password Check / temp password 123 for testing only
                    If UCase(.Range("C" & i).Value) = "TRUE" Then '<~~ Admin Check
                        CorrectLoginDetails = True
                        '
                        '~~> Do what you want
                        '
                        Me.AdminModeNewUserProfileLbl.Visible = True
                        Me.AdminModeNewUserProfile.Visible = True
                        Me.AdminModeLoginAddBtn.Caption = "Add"
                        Me.AdminModeNewUserProfile.Value = Environ("Username")
                        
                        wsAdmin.Visible = xlSheetVisible '<--- to remove after testing
                        wsAdmin.Activate '<--- to remove after testing
                    Else
                        '
                        '~~> Do what you want
                        '
                        wsAdmin.Visible = xlSheetVeryHidden
                        Me.AdminModeLoginAddBtn.Caption = "Login"
                        Me.AdminModeMessageLbl.ForeColor = vbRed
                        Me.AdminModeMessageLbl.Caption = "You do not have admin privileges!"
                    End If

                    Exit For
                End If
            End If
        Next i

        '~~> Incorrect Username/Password
        If CorrectLoginDetails = False Then
            If Not wsAdmin.Visible = xlSheetVeryHidden Then wsAdmin.Visible = xlSheetVeryHidden
            Me.AdminModeMessageLbl.ForeColor = vbRed
            Me.AdminModeMessageLbl.Caption = "Invalid username or password. Please try again!"
        End If
    End With
End Sub
Private Sub AddUserProfile()
    
    Dim lrow As Long
      
    lrow = wsAdmin.Range("E" & wsAdmin.Rows.Count).End(xlUp).Row
    UserList = wsAdmin.Range("E2", wsAdmin.Cells(Rows.Count, "E").End(xlUp))
    
        If wsAdmin.Range("E" & lrow).Offset(1, 0).Value = vbNullString Then
            If isUniqueUser(AdminModeNewUserProfile.Value) = True Then
                wsAdmin.Range("E" & lrow).Offset(1, 0).Value = Trim(Me.AdminModeNewUserProfile.Value)
            End If
        End If
    
End Sub
Function isUniqueUser(strIn As Variant) As Boolean

    isUniqueUser = IsError(Application.Match(strIn, UserList, 0))
End Function
Private Function EverythingFilledInAdminMode() As Boolean

    Dim ctl As MSForms.Control
    Dim txt As MSForms.TextBox
    Dim cmb As MSForms.ComboBox
    Dim AnythingIsMissing As Boolean
    
    EverythingFilledInAdminMode = True
    AnythingIsMissing = False
    
    For Each ctl In AdminModeFrame.Controls
        
        If TypeOf ctl Is MSForms.TextBox Then
            If Not ctl.Visible = False Then
                Set txt = ctl
                If txt.Value = vbNullString Then
                    txt.BackColor = rgbPink
                    Controls(ctl.Name & "Lbl").ForeColor = rgbRed
                    Me.AdminModeMessageLbl.Caption = "One or more fields are empty. Please try again."
                    
                    If Not AnythingIsMissing Then ctl.SetFocus
                        AnythingIsMissing = True
                        EverythingFilledInAdminMode = False
                    Else
                        AnythingIsMissing = False
                        EverythingFilledInAdminMode = True
                    End If
                        If EverythingFilledInAdminMode = True Then AdminModeMessageLbl.Caption = ""
            End If
        End If
    Next ctl
    
Set ctl = Nothing

End Function
Private Sub AdminModeCancelPressed_Click()

    Me.AdminModeAdminName.Value = ""
    Me.AdminModePassword.Value = ""
    Me.AdminModeNewUserProfile.Value = ""
    
    Me.AdminModeAdminNameLbl.ForeColor = Me.ForeColor
    Me.AdminModeAdminName.BackColor = rgbWhite
    Me.AdminModePasswordLbl.ForeColor = Me.ForeColor
    Me.AdminModePassword.BackColor = rgbWhite
    Me.AdminModeNewUserProfileLbl.ForeColor = Me.ForeColor
    Me.AdminModeNewUserProfile.BackColor = rgbWhite
    
    Me.AdminModeMessageLbl = ""
    Me.AdminModeFrame.Visible = False
    
    FormDataMode = AllBtnsEnabled
End Sub
Private Sub FillingInForm()

    Dim i As Long
    Dim WS As Worksheet
    Dim lrow As Long
    
    Dim ctl As MSForms.Control
    Dim lbl As MSForms.Label
    Dim cmb As MSForms.ComboBox
    Dim txtbox As MSForms.TextBox
    Dim optbtn As MSForms.OptionButton
    
    Set WS = ActiveSheet
    
With WS
    For i = 1 To ItemsListFrame.Controls.Count
        For Each ctl In ItemsListFrame.Controls
            If TypeName(ctl) = "Label" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set lbl = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        .Range("A" & i + 6).Offset(0, 0).Value = Me.OrderNo.Value
                        .Range("A" & i + 6).Offset(0, 1).Value = Format(Me.NextCollectionDate.Text, "dd/mm/yyyy")
                        .Range("A" & i + 6).Offset(0, 1).Value = Format(.Range("A" & i + 6).Offset(0, 1).Value, "dd/mm/yyyy")
                        '.Range("A" & i + 6).Offset(0, 8).Value = Me.DateReturnBy.Value
                        .Range("A" & i + 6).Offset(0, 8).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                        Controls("OrderLbl" & i).Enabled = True
                    End If
                End If
    
            ElseIf TypeName(ctl) = "ComboBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set cmb = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        Controls("Item" & i).Enabled = True
                    End If
                    If Controls("Item" & i).Value <> vbNullString Then
                        .Range("A" & i + 6).Offset(0, 2).Value = Controls("Item" & i).Text
                    End If
                End If
    
            ElseIf TypeName(ctl) = "TextBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set txtbox = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        .Range("A" & i + 6).Offset(0, 3).Value = Controls("Qty" & i).Value
                        .Range("A" & i + 6).Offset(0, 4).Value = Controls("UnitPrice" & i).Value
                        .Range("A" & i + 6).Offset(0, 5).Value = Controls("SubTotal" & i).Value
                        .Range("A" & i + 6).Offset(0, 7).Value = Controls("Comments" & i).Value
                        .Range("A" & i + 6).Offset(0, 8).Value = Format(Controls("ReturnDate" & i).Value, "dd/mm/yyyy")
                        Controls("Qty" & i).Enabled = True
                        Controls("UnitPrice" & i).Enabled = True
                        Controls("SubTotal" & i).Enabled = True
                        Controls("Comments" & i).Enabled = True
                        Controls("ReturnDate" & i).Enabled = True
                    End If
                End If
            ElseIf TypeName(ctl) = "OptionButton" Then
                If ctl.Tag = "GroupItem" & i Or ctl.Tag = "InOut" & i Then
                    Set optbtn = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        If Controls("OptionOut" & i).Value = True Then
                            .Range("A" & i + 6).Offset(0, 6).Value = "OUT"
                            .Range("A" & i + 6).Offset(0, 8).Value = Format(Me.DateReturnBy.Value, "dd/mm/yyyy")
                            .Range("A" & i + 6).Offset(0, 8).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                             Controls("ReturnDate" & i).Value = Format(.Range("A" & i + 6).Offset(0, 8).Value, "dd/mm/yyyy")
                        Else
                            .Range("A" & i + 6).Offset(0, 6).Value = "IN"
                            .Range("A" & i + 6).Offset(0, 8).Value = ""
                            Controls("ReturnDate" & i).Value = ""
                        End If
                        Controls("OptionIn" & i).Enabled = True
                        Controls("OptionOut" & i).Enabled = True
                    End If
                End If
            End If
        Next ctl
      Next i
End With
    
End Sub
Private Sub Item1_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item2_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item3_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item4_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item5_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item6_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item7_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item8_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item9_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item10_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item11_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item12_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub Item13_AfterUpdate()
    Call EnabledItemsToClean
End Sub
Private Sub EnabledItemsToClean()
    Dim ctl As MSForms.Control
    Dim lbl As MSForms.Label
    Dim cmb As MSForms.ComboBox
    Dim txtbox As MSForms.TextBox
    Dim optbtn As MSForms.OptionButton
    Dim i As Integer


    For i = 1 To 12
        For Each ctl In ItemsListFrame.Controls
            If TypeName(ctl) = "Label" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set lbl = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        Controls("OrderLbl" & i + 1).Enabled = True
                    End If
                End If

            ElseIf TypeName(ctl) = "ComboBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set cmb = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        Controls("Item" & i + 1).Enabled = True
                    End If
                End If

            ElseIf TypeName(ctl) = "TextBox" Then
                If ctl.Tag = "GroupItem" & i Then
                    Set txtbox = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        Controls("Qty" & i + 1).Enabled = True
                        Controls("UnitPrice" & i + 1).Enabled = True
                        Controls("SubTotal" & i + 1).Enabled = True
                        Controls("Comments" & i + 1).Enabled = True
                        Controls("ReturnDate" & i + 1).Enabled = True
                    End If
                End If
            ElseIf TypeName(ctl) = "OptionButton" Then
                If ctl.Tag = "GroupItem" & i Or ctl.Tag = "InOut" & i Then
                    Set optbtn = ctl
                    If Controls("Item" & i).Value <> vbNullString Then
                        Controls("OptionIn" & i + 1).Enabled = True
                        Controls("OptionOut" & i + 1).Enabled = True
                    End If
                End If
            End If
        Next ctl
      Next i
End Sub
Private Sub UnitPrice1_AfterUpdate()
'testing not complete
    'Call UpdateSubTotal
End Sub
Private Sub UpdateSubTotal()
'testing no complete

    Dim ctl As MSForms.Control
    Dim txtbox As MSForms.TextBox
    Dim i As Integer


        For Each ctl In ItemsListFrame.Controls
            If TypeName(ctl) = "TextBox" Then
                If ctl.Tag = "GroupItem" & i - 1 Then
                    Set txtbox = ctl
                    If Controls("Item" & i - 1).Value <> vbNullString Then
                        Controls("Qty" & i).Enabled = True
                        Controls("UnitPrice" & i).Enabled = True
                        Controls("SubTotal" & i).Enabled = True
                        Controls("Comments" & i).Enabled = True
                        Controls("ReturnDate" & i).Enabled = True
                    End If
                End If
            End If
        Next ctl
End Sub

Private Sub SaveBtnPressed_Click()

    Call FillingInForm
    ThisWorkbook.Save

End Sub
Private Sub CloseEntireFormBtnPressed_Click()

    'Main cancel button to unload the entire form
    Unload Me
End Sub
Private Sub UserForm_Terminate()

'To remove or comment after project is finished
'Check to see if workbook is acctualy closing down and not staying in the background
'If show records is on then close form and click cancel on save pop up excel will be in background -> need to fix it
'Be vary carefull when using the .Quit
    If Workbooks.Count > 1 Then
            ThisWorkbook.Save
            Windows(SmcDryCleaningForm).Close
        Else
            ThisWorkbook.Save
            Windows(SmcDryCleaningForm).Close
            'Application.Quit
        End If
End Sub

