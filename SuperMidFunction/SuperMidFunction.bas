Attribute VB_Name = "SuperMid"
Option Explicit
Public Function SuperMid(ByVal strMain As String, str1 As String, str2 As String, Optional reverse As Boolean) As String
'DESCRIPTION: Extract the portion of a string between the two substrings defined in str1 and str2.
'HOW TO USE: - Pass the argument your main string and the 2 strings you want to find in the main string.
' - This function will extract the values between the end of your first string and the beginning of your next string.
' - If the optional boolean "reverse" is true, an InStrRev search will occur to find the last instance of the substrings in your main string.
Dim i As Integer, j As Integer, temp As Variant
On Error GoTo errhandler:
If reverse = True Then
    i = InStrRev(strMain, str1)
    j = InStrRev(strMain, str2)
    If Abs(j - i) < Len(str1) Then j = InStrRev(strMain, str2, i)
    If i = j Then 'try to search 2nd half of string for unique match
        j = InStrRev(strMain, str2, i - 1)
    End If
Else
    i = InStr(1, strMain, str1)
    j = InStr(1, strMain, str2)
    If Abs(j - i) < Len(str1) Then j = InStr(i + Len(str1), strMain, str2)
    If i = j Then 'try to search 2nd half of string for unique match
        j = InStr(i + 1, strMain, str2)
    End If
End If
If i = 0 And j = 0 Then GoTo errhandler:
If j = 0 Then j = Len(strMain) + Len(str2) 'just to make it arbitrarily large
If i = 0 Then i = Len(strMain) + Len(str1) 'just to make it arbitrarily large
If i > j And j <> 0 Then 'swap order
    temp = j
    j = i
    i = temp
    temp = str2
    str2 = str1
    str1 = temp
End If
i = i + Len(str1)
SuperMid = Mid(strMain, i, j - i)
Exit Function
errhandler:
MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & "Aborting", , "Strings not found"
End
End Function

