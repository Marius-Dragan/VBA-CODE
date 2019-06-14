Attribute VB_Name = "RibbonSetup"
Option Explicit
'Created by Marius Dragan on 22/07/2018.
'Copyright © 2018. All rights reserved.
Private Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
'PURPOSE: Show/Hide buttons based on how many you need (False = Hide/False = Show)
'Module naming must not have the same name as the sub-procedure!!! This will throw an error and macro will not run

Select Case control.ID
  Case "GroupA": MakeVisible = False
  Case "aButton01": MakeVisible = False
  Case "aButton02": MakeVisible = False
  Case "aButton03": MakeVisible = False
  Case "aButton04": MakeVisible = False
  Case "aButton05": MakeVisible = False
  Case "aButton06": MakeVisible = False
  Case "aButton07": MakeVisible = False
  Case "aButton08": MakeVisible = False
  Case "aButton09": MakeVisible = False
  Case "aButton10": MakeVisible = False
  
  Case "GroupB": MakeVisible = False
  Case "bButton01": MakeVisible = False
  Case "bButton02": MakeVisible = False
  Case "bButton03": MakeVisible = False
  Case "bButton04": MakeVisible = False
  Case "bButton05": MakeVisible = False
  Case "bButton06": MakeVisible = False
  Case "bButton07": MakeVisible = False
  Case "bButton08": MakeVisible = False
  Case "bButton09": MakeVisible = False
  Case "bButton10": MakeVisible = False
  
  Case "GroupC": MakeVisible = False
  Case "cButton01": MakeVisible = False
  Case "cButton02": MakeVisible = False
  Case "cButton03": MakeVisible = False
  Case "cButton04": MakeVisible = False
  Case "cButton05": MakeVisible = False
  Case "cButton06": MakeVisible = False
  Case "cButton07": MakeVisible = False
  Case "cButton08": MakeVisible = False
  Case "cButton09": MakeVisible = False
  Case "cButton10": MakeVisible = False
  
  Case "GroupD": MakeVisible = True
  Case "dButton01": MakeVisible = True
  Case "dButton02": MakeVisible = True
  Case "dButton03": MakeVisible = True
  Case "dButton04": MakeVisible = True
  Case "dButton05": MakeVisible = True
  Case "dButton06": MakeVisible = True
  Case "dButton07": MakeVisible = True
  Case "dButton08": MakeVisible = True
  Case "dButton09": MakeVisible = True
  Case "dButton10": MakeVisible = True
  
  Case "GroupE": MakeVisible = False
  Case "eButton01": MakeVisible = False
  Case "eButton02": MakeVisible = False
  Case "eButton03": MakeVisible = False
  Case "eButton04": MakeVisible = False
  Case "eButton05": MakeVisible = False
  Case "eButton06": MakeVisible = False
  Case "eButton07": MakeVisible = False
  Case "eButton08": MakeVisible = False
  Case "eButton09": MakeVisible = False
  Case "eButton10": MakeVisible = False
  
  Case "GroupF": MakeVisible = True
  Case "fButton01": MakeVisible = True
  Case "fButton02": MakeVisible = True
  Case "fButton03": MakeVisible = True
  Case "fButton04": MakeVisible = True
  Case "fButton05": MakeVisible = True
  Case "fButton06": MakeVisible = True
  Case "fButton07": MakeVisible = True
  Case "fButton08": MakeVisible = True
  Case "fButton09": MakeVisible = True
  Case "fButton10": MakeVisible = True
  
End Select

End Sub

Private Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling)
'PURPOSE: Determine the text to go along with your Tab, Groups, and Buttons

Select Case control.ID
  
  Case "CustomTab": Labeling = "Marius"
  
  Case "GroupA": Labeling = "Inventory Reports"
  Case "aButton01": Labeling = "Button"
  Case "aButton02": Labeling = "Button"
  Case "aButton03": Labeling = "Button"
  Case "aButton04": Labeling = "Button"
  Case "aButton05": Labeling = "Button"
  Case "aButton06": Labeling = "Button"
  Case "aButton07": Labeling = "Button"
  Case "aButton08": Labeling = "Button"
  Case "aButton09": Labeling = "Button"
  Case "aButton10": Labeling = "Button"
  
  Case "GroupB": Labeling = "Markdown Reports"
  Case "bButton01": Labeling = "Button"
  Case "bButton02": Labeling = "Button"
  Case "bButton03": Labeling = "Button"
  Case "bButton04": Labeling = "Button"
  Case "bButton05": Labeling = "Button"
  Case "bButton06": Labeling = "Button"
  Case "bButton07": Labeling = "Button"
  Case "bButton08": Labeling = "Button"
  Case "bButton09": Labeling = "Button"
  Case "bButton10": Labeling = "Button"
  
  Case "GroupC": Labeling = "RTV Reports"
  Case "cButton01": Labeling = "Button"
  Case "cButton02": Labeling = "Button"
  Case "cButton03": Labeling = "Button"
  Case "cButton04": Labeling = "Button"
  Case "cButton05": Labeling = "Button"
  Case "cButton06": Labeling = "Button"
  Case "cButton07": Labeling = "Button"
  Case "cButton08": Labeling = "Button"
  Case "cButton09": Labeling = "Button"
  Case "cButton10": Labeling = "Button"
  
  Case "GroupD": Labeling = "Utilities"
  Case "dButton01": Labeling = "Advanced Multiple Criteria Filter"
  Case "dButton02": Labeling = "Reset Filters"
  Case "dButton03": Labeling = "Comparing Two Lists"
  Case "dButton04": Labeling = "Auto Fit Cells"
  Case "dButton05": Labeling = "Cross-Highlight cells"
  Case "dButton06": Labeling = "Consolidating Multiple WS"
  Case "dButton07": Labeling = "Consolidating Multiple WB"
  Case "dButton08": Labeling = "Clear Only Values"
  Case "dButton09": Labeling = "Version Rename"
  Case "dButton10": Labeling = "Copy Each WS To New WB"
  
  Case "GroupE": Labeling = "Utilities"
  Case "eButton01": Labeling = "Button"
  Case "eButton02": Labeling = "Button"
  Case "eButton03": Labeling = "Button"
  Case "eButton04": Labeling = "Button"
  Case "eButton05": Labeling = "Button"
  Case "eButton06": Labeling = "Button"
  Case "eButton07": Labeling = "Button"
  Case "eButton08": Labeling = "Button"
  Case "eButton09": Labeling = "Button"
  Case "eButton10": Labeling = "Button"
  
  Case "GroupF": Labeling = "Worksheet"
  Case "fButton01": Labeling = "About"
  Case "fButton02": Labeling = "Toggle Case Text"
  Case "fButton03": Labeling = "Proper Case Text"
  Case "fButton04": Labeling = "Upper Case Text"
  Case "fButton05": Labeling = "Lower Case Text"
  Case "fButton06": Labeling = "Count Non Blank Cells"
  Case "fButton07": Labeling = "Split Style Fabric Colour Size"
  Case "fButton08": Labeling = "Trim Non Printable Chr"
  Case "fButton09": Labeling = "Sort active worksheets"
  Case "fButton10": Labeling = "Edit Printing Properties"
  
End Select
   
End Sub

Private Sub GetImage(control As IRibbonControl, ByRef RibbonImage)
'PURPOSE: Tell each button which image to load from the Microsoft Icon Library
'TIPS: Image names are case sensitive, if image does not appear in ribbon after re-starting Excel, the image name is incorrect

Select Case control.ID
  
  Case "aButton01": RibbonImage = "AddContentType"
  Case "aButton02": RibbonImage = "AutoFormatNow"
  Case "aButton03": RibbonImage = "AddressBook"
  Case "aButton04": RibbonImage = "ImportTextFile"
  Case "aButton05": RibbonImage = "AddContentType"
  Case "aButton06": RibbonImage = "ObjectPictureFill"
  Case "aButton07": RibbonImage = "ObjectPictureFill"
  Case "aButton08": RibbonImage = "ObjectPictureFill"
  Case "aButton09": RibbonImage = "ObjectPictureFill"
  Case "aButton10": RibbonImage = "ObjectPictureFill"
  
  Case "bButton01": RibbonImage = "AddContentType"
  Case "bButton02": RibbonImage = "AutoFormatDialog"
  Case "bButton03": RibbonImage = "NewFormGallery"
  Case "bButton04": RibbonImage = "ObjectPictureFill"
  Case "bButton05": RibbonImage = "ObjectPictureFill"
  Case "bButton06": RibbonImage = "ObjectPictureFill"
  Case "bButton07": RibbonImage = "ObjectPictureFill"
  Case "bButton08": RibbonImage = "ObjectPictureFill"
  Case "bButton09": RibbonImage = "ObjectPictureFill"
  Case "bButton10": RibbonImage = "ObjectPictureFill"
  
  Case "cButton01": RibbonImage = "ResourceDetailsDisplay"
  Case "cButton02": RibbonImage = "AddContentType"
  Case "cButton03": RibbonImage = "AddContentType"
  Case "cButton04": RibbonImage = "ObjectPictureFill"
  Case "cButton05": RibbonImage = "ObjectPictureFill"
  Case "cButton06": RibbonImage = "ObjectPictureFill"
  Case "cButton07": RibbonImage = "ObjectPictureFill"
  Case "cButton08": RibbonImage = "ObjectPictureFill"
  Case "cButton09": RibbonImage = "ObjectPictureFill"
  Case "cButton10": RibbonImage = "ObjectPictureFill"
  
  Case "dButton01": RibbonImage = "AutoFilterProject"
  Case "dButton02": RibbonImage = "CancelRequest"
  Case "dButton03": RibbonImage = "AddContentType"
  Case "dButton04": RibbonImage = "CloseWeb"
  Case "dButton05": RibbonImage = "ChartInsert"
  Case "dButton06": RibbonImage = "ObjectPictureFill"
  Case "dButton07": RibbonImage = "ObjectPictureFill"
  Case "dButton08": RibbonImage = "ObjectPictureFill"
  Case "dButton09": RibbonImage = "ObjectPictureFill"
  Case "dButton10": RibbonImage = "ObjectPictureFill"
  
  Case "eButton01": RibbonImage = "PageSetupPageDialog"
  Case "eButton02": RibbonImage = "TripaneViewMode"
  Case "eButton03": RibbonImage = "TaskInformation"
  Case "eButton04": RibbonImage = "SharePointListsWorkOffline"
  Case "eButton05": RibbonImage = "ObjectPictureFill"
  Case "eButton06": RibbonImage = "ObjectPictureFill"
  Case "eButton07": RibbonImage = "ObjectPictureFill"
  Case "eButton08": RibbonImage = "ObjectPictureFill"
  Case "eButton09": RibbonImage = "ObjectPictureFill"
  Case "eButton10": RibbonImage = "ObjectPictureFill"
  
  Case "fButton01": RibbonImage = "Info"
  Case "fButton02": RibbonImage = "PhoneticGuideSettings"
  Case "fButton03": RibbonImage = "MessageFormatRichText"
  Case "fButton04": RibbonImage = "SmartArtIncreaseFontSize"
  Case "fButton05": RibbonImage = "StylisticAlternate"
  Case "fButton06": RibbonImage = "SmartArtTextPane"
  Case "fButton07": RibbonImage = "ObjectsAlignRight"
  Case "fButton08": RibbonImage = "SparklineCustomWeight"
  Case "fButton09": RibbonImage = "SortDialog"
  Case "fButton10": RibbonImage = "PageSettings"
  
End Select

End Sub

Private Sub GetSize(control As IRibbonControl, ByRef size)
'PURPOSE: Determine if the button size is large or small

Const Large As Integer = 1
Const Small As Integer = 0

Select Case control.ID
    
  Case "aButton01": size = Small
  Case "aButton02": size = Small
  Case "aButton03": size = Small
  Case "aButton04": size = Small
  Case "aButton05": size = Small
  Case "aButton06": size = Small
  Case "aButton07": size = Small
  Case "aButton08": size = Small
  Case "aButton09": size = Small
  Case "aButton10": size = Small
  
  Case "bButton01": size = Small
  Case "bButton02": size = Small
  Case "bButton03": size = Small
  Case "bButton04": size = Small
  Case "bButton05": size = Small
  Case "bButton06": size = Small
  Case "bButton07": size = Small
  Case "bButton08": size = Small
  Case "bButton09": size = Small
  Case "bButton10": size = Small
  
  Case "cButton01": size = Small
  Case "cButton02": size = Small
  Case "cButton03": size = Small
  Case "cButton04": size = Small
  Case "cButton05": size = Small
  Case "cButton06": size = Small
  Case "cButton07": size = Small
  Case "cButton08": size = Small
  Case "cButton09": size = Small
  Case "cButton10": size = Small
  
  Case "dButton01": size = Large
  Case "dButton02": size = Small
  Case "dButton03": size = Small
  Case "dButton04": size = Small
  Case "dButton05": size = Small
  Case "dButton06": size = Small
  Case "dButton07": size = Small
  Case "dButton08": size = Small
  Case "dButton09": size = Small
  Case "dButton10": size = Small
  
  Case "eButton01": size = Small
  Case "eButton02": size = Small
  Case "eButton03": size = Small
  Case "eButton04": size = Small
  Case "eButton05": size = Small
  Case "eButton06": size = Small
  Case "eButton07": size = Small
  Case "eButton08": size = Small
  Case "eButton09": size = Small
  Case "eButton10": size = Small
  
  Case "fButton01": size = Large
  Case "fButton02": size = Small
  Case "fButton03": size = Small
  Case "fButton04": size = Small
  Case "fButton05": size = Small
  Case "fButton06": size = Small
  Case "fButton07": size = Small
  Case "fButton08": size = Small
  Case "fButton09": size = Small
  Case "fButton10": size = Small
  
End Select

End Sub

Private Sub RunMacro(control As IRibbonControl)
'PURPOSE: Tell each button which macro subroutine to run when clicked

Select Case control.ID
  
  Case "aButton01": Application.Run "TestingButtons"
  Case "aButton02": Application.Run "TestingButtons"
  Case "aButton03": Application.Run "TestingButtons"
  Case "aButton04": Application.Run "TestingButtons"
  Case "aButton05": Application.Run "TestingButtons"
  Case "aButton06": Application.Run "TestingButtons"
  Case "aButton07": Application.Run "TestingButtons"
  Case "aButton08": Application.Run "TestingButtons"
  Case "aButton09": Application.Run "TestingButtons"
  Case "aButton10": Application.Run "TestingButtons"
  
  Case "bButton01": Application.Run "TestingButtons"
  Case "bButton02": Application.Run "TestingButtons"
  Case "bButton03": Application.Run "TestingButtons"
  Case "bButton04": Application.Run "TestingButtons"
  Case "bButton05": Application.Run "TestingButtons"
  Case "bButton06": Application.Run "TestingButtons"
  Case "bButton07": Application.Run "TestingButtons"
  Case "bButton08": Application.Run "TestingButtons"
  Case "bButton09": Application.Run "TestingButtons"
  Case "bButton10": Application.Run "TestingButtons"
  
  Case "cButton01": Application.Run "TestingButtons"
  Case "cButton02": Application.Run "TestingButtons"
  Case "cButton03": Application.Run "TestingButtons"
  Case "cButton04": Application.Run "TestingButtons"
  Case "cButton05": Application.Run "TestingButtons"
  Case "cButton06": Application.Run "TestingButtons"
  Case "cButton07": Application.Run "TestingButtons"
  Case "cButton08": Application.Run "TestingButtons"
  Case "cButton09": Application.Run "TestingButtons"
  Case "cButton10": Application.Run "TestingButtons"
  
  Case "dButton01": Application.Run "AdvancedMultipleCriteriaFilter"
  Case "dButton02": Application.Run "ResetFilters"
  Case "dButton03": Application.Run "ComparingTwoLists"
  Case "dButton04": Application.Run "AutoFit"
  Case "dButton05": Application.Run "AutoHighlightSelectedCells"
  Case "dButton06": Application.Run "ConsolidatingMultipleWS"
  Case "dButton07": Application.Run "ConsolidatingMultipleWB"
  Case "dButton08": Application.Run "ClearOnlyValues"
  Case "dButton09": Application.Run "VersionRename"
  Case "dButton10": Application.Run "CopyEachWSToNewWB"
  
  Case "eButton01": Application.Run "TestingButtons"
  Case "eButton02": Application.Run "TestingButtons"
  Case "eButton03": Application.Run "TestingButtons"
  Case "eButton04": Application.Run "TestingButtons"
  Case "eButton05": Application.Run "TestingButtons"
  Case "eButton06": Application.Run "ConsolidatingMultipleWB"
  Case "eButton07": Application.Run "TestingButtons"
  Case "eButton08": Application.Run "TestingButtons"
  Case "eButton09": Application.Run "TestingButtons"
  Case "eButton10": Application.Run "TestingButtons"
  
  Case "fButton01": Application.Run "VBA_Version"
  Case "fButton02": Application.Run "ToggleCaseMacro"
  Case "fButton03": Application.Run "ProperMacro"
  Case "fButton04": Application.Run "UpperMacro"
  Case "fButton05": Application.Run "LowerMacro"
  Case "fButton06": Application.Run "CountNonBlankCells"
  Case "fButton07": Application.Run "SplitStyleFabricColourSize"
  Case "fButton08": Application.Run "TrimCellsArrayMethodSelection"
  Case "fButton09": Application.Run "SortActiveWorksheets"
  Case "fButton10": Application.Run "EditPrintingProperties"

 End Select
    
End Sub

Private Sub GetScreentip(control As IRibbonControl, ByRef Screentip)
'PURPOSE: Display a specific macro description when the mouse hovers over a button

Select Case control.ID
  
  Case "aButton01": Screentip = "Description"
  Case "aButton02": Screentip = "Description"
  Case "aButton03": Screentip = "Description"
  Case "aButton04": Screentip = "Description"
  Case "aButton05": Screentip = "Description"
  Case "aButton06": Screentip = "Description"
  Case "aButton07": Screentip = "Description"
  Case "aButton08": Screentip = "Description"
  Case "aButton09": Screentip = "Description"
  Case "aButton10": Screentip = "Description"
  
  Case "bButton01": Screentip = "Description"
  Case "bButton02": Screentip = "Description"
  Case "bButton03": Screentip = "Description"
  Case "bButton04": Screentip = "Description"
  Case "bButton05": Screentip = "Description"
  Case "bButton06": Screentip = "Description"
  Case "bButton07": Screentip = "Description"
  Case "bButton08": Screentip = "Description"
  Case "bButton09": Screentip = "Description"
  Case "bButton10": Screentip = "Description"
  
  Case "cButton01": Screentip = "Description"
  Case "cButton02": Screentip = "Description"
  Case "cButton03": Screentip = "Description"
  Case "cButton04": Screentip = "Description"
  Case "cButton05": Screentip = "Description"
  Case "cButton06": Screentip = "Description"
  Case "cButton07": Screentip = "Description"
  Case "cButton08": Screentip = "Description"
  Case "cButton09": Screentip = "Description"
  Case "cButton10": Screentip = "Description"
  
  Case "dButton01": Screentip = "Advanced multiple criteria filter from user selection"
  Case "dButton02": Screentip = "Reset all filters on worksheet"
  Case "dButton03": Screentip = "Comparing two lists and returning any result from user input"
  Case "dButton04": Screentip = "Autofit all cells"
  Case "dButton05": Screentip = "Auto-Highlight selected cells"
  Case "dButton06": Screentip = "Consolidating multiple WS"
  Case "dButton07": Screentip = "Consolidating multiple WB"
  Case "dButton08": Screentip = "Clear only values"
  Case "dButton09": Screentip = "Rename All Files In Folder with _V + No"
  Case "dButton10": Screentip = "Copy Each WS To New WB"

  Case "eButton01": Screentip = "Description"
  Case "eButton02": Screentip = "Description"
  Case "eButton03": Screentip = "Description"
  Case "eButton04": Screentip = "Description"
  Case "eButton05": Screentip = "Auto fit cells from current region"
  Case "eButton06": Screentip = "Description"
  Case "eButton07": Screentip = "Description"
  Case "eButton08": Screentip = "Description"
  Case "eButton09": Screentip = "Description"
  Case "eButton10": Screentip = "Description"
  
  Case "fButton01": Screentip = "Creat a new worksheet with all macro details"
  Case "fButton02": Screentip = "Toggle case from user selection"
  Case "fButton03": Screentip = "Proper case from user selection"
  Case "fButton04": Screentip = "Upper case from user selection"
  Case "fButton05": Screentip = "Lower case from user selection"
  Case "fButton06": Screentip = "Count non blank cells from user selection"
  Case "fButton07": Screentip = "Split style fabric colour size from user selection"
  Case "fButton08": Screentip = "Trims non printable characters from user selection"
  Case "fButton09": Screentip = "Sorts active worksheets in a workbook"
  Case "fButton10": Screentip = "Edit printing properties landscape orientation"
  
End Select

End Sub

Private Sub TestingButtons()
MsgBox "Button works successfully with no errors!"
End Sub

