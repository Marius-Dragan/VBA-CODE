Attribute VB_Name = "removeSpace"
Sub removeSpaceV2()
'
' Working version with no errors.

    Dim rngRemoveSpace As Range
    Dim CellChecker As Range
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
     'On Error GoTo ErrorHandler
     
  Set rngRemoveSpace = Intersect(ActiveSheet.UsedRange, Selection)
    rngRemoveSpace.Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
    
    For Each CellChecker In rngRemoveSpace.Cells
        CellChecker.value = Application.Trim(CellChecker.value)
        CellChecker.value = Application.Clean(CellChecker.value)
        
    Next CellChecker
    
Application.ScreenUpdating = True

    Set rngRemoveSpace = Nothing
'ErrorHandler:
   'Debug.Print "Error number: " & Err.Number & " " & Err.Description
'        MsgBox "Sorry, an error occured." & vbCrLf & Err.Description, vbCritical, "Error!"
    
End Sub
