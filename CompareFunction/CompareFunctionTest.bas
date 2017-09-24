Attribute VB_Name = "CompareFunctionTest"
Function compare(ByVal Cell1 As Range, ByVal Cell2 As Range, Optional CaseSensitive As Variant, Optional delta As Variant, Optional MatchString As Variant)
''***************************************************************************
''***DESCRIPTION: Compares Cell1 to Cell2 and if identical, returns "-" by *
''*** default but a different optional match string can be given. *
''*** If cells are different, the output will either be "FALSE" *
''*** or will optionally show the delta between the values if *
''*** numeric. *
''***INPUT: Cell1 - First cell to compare. *
''*** Cell2 - Cell to compare against Cell1. *
''*** CaseSensitive - Optional boolean that if set to TRUE, will *
''*** perform a case-sensitive comparison of the *
''*** two entered cells. Default is TRUE. *
''*** delta - Optional boolean that if set to TRUE, will display *
''*** the delta between Cell1 and Cell2. *
''*** MatchString - Optional string the user can choose to display *
''*** when Cell1 and Cell2 match. Default is "-" *
''***OUTPUT: The output will be "-", a custom string or a delta if the *
''*** cells match and will be "FALSE" if the cells do not match. *
''***EXAMPLES: =compare(A1,B1,FALSE,TRUE,"match") *
''*** =compare(A1,B1) *
''******************************************************************************

''------------------------------------------------------------------------------
''I. Declare variables
''------------------------------------------------------------------------------
Dim strMatch As String 'string to display if Cell1 and Cell2 match

''------------------------------------------------------------------------------
''II. Error checking
''------------------------------------------------------------------------------
''Error 0 - catch all error
On Error GoTo CompareError:

''Error 1 - MatchString is invalid
If IsMissing(MatchString) = False Then
    If IsError(CStr(MatchString)) Then
        compare = "Invalid Match String"
        Exit Function
    End If
End If

''Error 2 - Cell1 contains more than 1 cell
If IsArray(Cell1) = True Then
    If Cell1.Count <> 1 Then
        compare = "Too many cells in variable Cell1."
        Exit Function
    End If
End If

''Error 3 - Cell2 contains more than 1 cell
If IsArray(Cell2) = True Then
    If Cell2.Count <> 0 Then
        compare = "Too many cells in variable Cell2."
        Exit Function
    End If
End If

''Error 4 - delta is not a boolean
If IsMissing(delta) = False Then
    If delta <> CBool(True) And delta <> CBool(False) Then
        compare = "Delta flag must be a boolean (TRUE or FALSE)."
        Exit Function
    End If
End If

''Error 5 - CaseSensitive is not a boolean
If IsMissing(CaseSensitive) = False Then
    If CaseSensitive <> CBool(True) And CaseSensitive <> CBool(False) Then
        compare = "CaseSensitive flag must be a boolean (TRUE or FALSE)."
        Exit Function
    End If
End If

''------------------------------------------------------------------------------
''III. Initialize Variables
''------------------------------------------------------------------------------
If IsMissing(CaseSensitive) Then
    CaseSensitive = CBool(True)
ElseIf CaseSensitive = False Then
    CaseSensitive = CBool(False)
Else
    CaseSensitive = CBool(True)
End If

If IsMissing(MatchString) Then
    strMatch = "-"
Else
    strMatch = CStr(MatchString)
End If

If IsMissing(delta) Then
    delta = CBool(False)
ElseIf delta = False Then
    delta = CBool(False)
Else
    delta = CBool(True)
End If

''------------------------------------------------------------------------------
''IV. Check for matches
''------------------------------------------------------------------------------
If Cell1 = Cell2 Then
    compare = strMatch
ElseIf CaseSensitive = False Then
    If UCase(Cell1) = UCase(Cell2) Then
        compare = strMatch
    ElseIf delta = True And IsNumeric(Cell1) And IsNumeric(Cell2) Then
        compare = Cell1 - Cell2
    Else
        compare = CBool(False)
    End If
ElseIf Cell1 <> Cell2 And delta = True Then
    If IsNumeric(Cell1) And IsNumeric(Cell2) Then
        'No case sensitive check because if not numeric, doesn't matter.
        compare = Cell1 - Cell2
    Else
        compare = CBool(False)
    End If
Else
    compare = CBool(False)
End If
Exit Function

''------------------------------------------------------------------------------
''V. Final Error Handling
''------------------------------------------------------------------------------
CompareError:
    compare = "Error Encountered: " & Err.Number & ", " & Err.Description
End Function
