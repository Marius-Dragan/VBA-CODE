Attribute VB_Name = "splitStyleFabricColourSize"
Option Explicit
Sub splitStyleFabricColourSizeV3()
Attribute splitStyleFabricColourSizeV3.VB_Description = "Split Style/Fabric/Colour/Size into 3-4 columns across"
Attribute splitStyleFabricColourSizeV3.VB_ProcData.VB_Invoke_Func = "T\n14"
'
'Please note you need to add a references to Microsoft VBScript Regular Expession 5.5
'Please note: you need to import poductCode class in order for this to work
'

     Dim wsSrc As Worksheet, wsRes As Worksheet
    Dim vSrc As Variant, vRes As Variant, rRes As Range
    Dim RE As Object, MC As Object

    Const sPat As String = "^(.{6})\s*(.{5})\s*(.{4})(?:.*1/(\S+))?"
        'Group 1 = style
        'Group 2 = fabric
        'Group 3 = colour
        'Group 4 = size
    Dim colF As Collection, pC As productCode
    Dim i As Long
    Dim S As String
    Dim V As Variant

'Set source and results worksheets and ranges
Set wsSrc = ActiveSheet
Set wsRes = ActiveSheet
    Set rRes = wsRes.Application.Selection

'Read source data into array
vSrc = Selection.Resize(columnsize:=4)

'Initialize the Collection object
Set colF = New Collection

'Initialize the Regex Object
Set RE = CreateObject("vbscript.regexp")
With RE
    .Global = False
    .MultiLine = True
    .Pattern = sPat

'Test for single cell
If Not IsArray(vSrc) Then
    V = vSrc
    ReDim vSrc(1 To 1, 1 To 1)
    vSrc(1, 1) = V
End If

    'iterate through the list
For i = 1 To UBound(vSrc, 1)
    S = vSrc(i, 1)
    Set pC = New productCode
    If .Test(S) = True Then
        Set MC = .Execute(S)
        With MC(0)
            pC.Style = .submatches(0)
            pC.Fabric = .submatches(1)
            pC.Colour = .submatches(2)
            pC.Size = .submatches(3)
        End With
         ElseIf .Test(vSrc(i, 1) & vSrc(i, 2) & vSrc(i, 3)) = False Then
        pC.Style = S
    Else
        pC.Style = vSrc(i, 1)
        pC.Fabric = vSrc(i, 2)
        pC.Colour = vSrc(i, 3)
        pC.Size = vSrc(i, 4)
    End If
    colF.Add pC
Next i
End With

'create results array
'Exit if not results
If colF.Count = 0 Then Exit Sub

ReDim vRes(1 To colF.Count, 1 To 4)

'Populate the rest
i = 0
For Each V In colF
    i = i + 1
    With V
        vRes(i, 1) = .Style
        vRes(i, 2) = .Fabric
        vRes(i, 3) = .Colour
        vRes(i, 4) = .Size

    End With
Next V

'Write the results
Set rRes = rRes.Resize(UBound(vRes, 1), UBound(vRes, 2))
    rRes.value = vRes

End Sub
