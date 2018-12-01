Attribute VB_Name = "SendEmail"
Option Explicit
Dim titleName As String
Dim firstName As String
Dim lastName As String
Dim fullName As String
Dim clientEmail As String
Dim ccEmail As String
Dim bccEmail As String
Dim emailMessage As String
 
Sub GenerateInfoToCreateEmail()
 
    Dim WS As Worksheet
    Dim lRow As Long
    Dim cRow As Long
    
    Set WS = ActiveSheet
 
    With WS
        lRow = .Range("E" & .Rows.Count).End(xlUp).Row
        Application.ScreenUpdating = False
        For cRow = 2 To lRow
            If Not .Range("L" & cRow).value = "" Then
                titleName = .Range("D" & cRow).value
                firstName = .Range("E" & cRow).value
                lastName = .Range("F" & cRow).value
                fullName = firstName & " " & lastName
                clientEmail = .Range("L" & cRow).value
                
                Call SendEmail
                
                .Range("Y" & cRow).value = "Yes"
                .Range("Y" & cRow).Font.Color = vbGreen
            
            Else
                .Range("Y" & cRow).value = "No"
                .Range("Y" & cRow).Font.Color = vbRed
            End If
        Next cRow
    End With
 
    Application.ScreenUpdating = True
    
    MsgBox "Process completed!", vbInformation
    
End Sub
Private Sub SendEmail()
 
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim sigString As String
    Dim Signature As String
    Dim insertPhoto As String
    Dim photoSize As String
 
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)
    
    'Change only Mysig.htm to the name of your signature
    sigString = Environ("appdata") & _
                "\Microsoft\Signatures\Marius.htm"
 
    If Dir(sigString) <> "" Then
        Signature = GetBoiler(sigString)
    Else
        Signature = ""
    End If
    
    insertPhoto = "C:\Users\marius.dragan\Desktop\Presale.jpg" 'Picture path
    photoSize = "<img src=""cid:Presale.jpg""height=400 width=400>" 'Change image name here
    
    emailMessage = "<BODY style=font-size:11pt;font-family:Calibri>Dear " & titleName & " " & fullName & "," & _
                    "<p>I hope my email will find you very well." & _
                    "<p>Our <strong>sales preview</strong> starts on Thursday the 22nd until Sunday the 25th of November." & _
                    "<p>I look forward to welcoming you into the store to shop on preview.<p>" & _
                    "<p> It really is the perfect opportunity to get some fabulous pieces for the fast approaching festive season." & _
                    "<p>Please feel free to contact me and book an appointment." & _
                    "<p>I look forward to seeing you then." & _
                    "<p>" & photoSize & _
                    "<p>Kind Regards," & _
                    "<br>" & _
                    "<br><strong>Marius Dragan</strong>" & _
                    "<br>Stock Controller" & _
                    "<p>"

    
    With outlookMail
        .To = clientEmail
        .CC = ""
        .BCC = ""
        .Subject = "PRIVATE SALE"
        .BodyFormat = 2
        .Attachments.Add insertPhoto, 1, 0
        .HTMLBody = emailMessage & Signature 'Including photo insert and signature
        '.HTMLBody = emailMessage & Signature 'Only signature
        .Importance = 2
        .ReadReceiptRequested = True
        .Display
        '.Send 'this will send the email without review
        
    End With
    
    Set outlookApp = Nothing
    Set outlookMail = Nothing
End Sub
Function GetBoiler(ByVal sFile As String) As String
 
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function


