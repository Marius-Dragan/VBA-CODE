Attribute VB_Name = "ClearAllData"
Option Explicit
Sub ClearAllData()
'
'
Dim dataToClear As Long

     OptimizeVBA True

    With Worksheets(2)
        dataToClear = Application.Max(4, _
                                    .Cells(.Rows.Count, "K").End(xlUp).Row, _
                                    .Cells(.Rows.Count, "L").End(xlUp).Row, _
                                    .Cells(.Rows.Count, "M").End(xlUp).Row)
                                    .Cells(4, "H").Resize(dataToClear - 4 + 1, 7).ClearContents

      End With
    ActiveSheet.PageSetup.PrintArea = ""

  OptimizeVBA False
End Sub
