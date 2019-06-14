Attribute VB_Name = "RenamingFiles"
Option Explicit
Sub VersionRename()

Dim SelectedFolder As FileDialog
Dim T_Str As String
Dim FSO As Object
Dim RenamingFolder As Object, SubFolder As Object
Dim T_Name As String
    
Set SelectedFolder = Application.FileDialog(msoFileDialogFolderPicker)
    SelectedFolder.Title = "Select folder:"
    SelectedFolder.ButtonName = "Select Folder"
    If SelectedFolder.Show = -1 Then
        T_Str = SelectedFolder.SelectedItems(1)
    Else
        'MsgBox "Cancelled by user.", vbInformation
    Set SelectedFolder = Nothing
    Exit Sub
    End If
    
    Set SelectedFolder = Nothing
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set RenamingFolder = FSO.GetFolder(T_Str)
        File_Renamer RenamingFolder
        
    For Each SubFolder In RenamingFolder.SubFolders
        File_Renamer SubFolder
    Next
    
    Set SubFolder = Nothing
    Set RenamingFolder = Nothing
    Set FSO = Nothing
    
    MsgBox "Process completed!", vbInformation, Title:="Renaming Files"


End Sub
Private Sub File_Renamer(Folder As Object)

Dim File As Object
Dim T_Str As String
Dim T_Name As String
Dim PreVersionID As Variant 'use variant when different types cone be return vrom inputbox
Dim NextVersionID As Variant 'use variant when different types cone be return vrom inputbox
Dim StringReplace As String

    PreVersionID = Application.InputBox("Input 1 if no version number otherwise input existing version number:", Type:=1)
    If PreVersionID = False Then Exit Sub
    NextVersionID = Application.InputBox("Input your next version number:", Type:=1)
    If NextVersionID = False Then Exit Sub


    T_Str = Format("_V" & NextVersionID)
    
    For Each File In Folder.Files
        T_Name = File.Name
        'Debug.Print T_Name
        If NextVersionID > 1 Then
            StringReplace = Replace(T_Name, "_V" & PreVersionID, "", 1, 3)
            'Debug.Print StringReplace
            File.Name = Left(StringReplace, InStrRev(StringReplace, ".") - 1) & T_Str & Right(StringReplace, Len(StringReplace) - (InStrRev(StringReplace, ".") - 1))
        Else
            File.Name = Left(T_Name, InStrRev(T_Name, ".") - 1) & T_Str & Right(T_Name, Len(T_Name) - (InStrRev(T_Name, ".") - 1))
        End If
    Next
End Sub


