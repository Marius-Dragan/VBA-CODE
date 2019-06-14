Attribute VB_Name = "CalcCheckDigit"
Option Explicit
Function CalculateCheckDigitUPC(value)

Dim lenghtVal As Long
Dim factor As Long
Dim sum As Long
Dim index As Long
    lenghtVal = Len(value)
    factor = 3
    sum = 0
    For index = lenghtVal To 1 Step -1
        sum = sum + (CInt(Mid(value, index, 1)) * factor)
        factor = 4 - factor
    Next
    CalculateCheckDigitUPC = ((1000 - sum) Mod 10)
End Function
Sub GenerateCheckDigit()

Dim cell As Variant
Dim currentValue As String

    If TypeName(Selection) <> "Range" Then Exit Sub
        For Each cell In Selection
            currentValue = cell.value
            cell.Offset(0, 2).value = currentValue & CalculateCheckDigitUPC(currentValue)
        Next
End Sub
