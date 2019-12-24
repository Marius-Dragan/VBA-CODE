Attribute VB_Name = "UserM3ProfileChange"
Option Explicit
Sub UpdateUserProfile()

Dim USID, CONO, DIVI, FACI, WHLO
Dim strUsername, strPassword As String
Dim apiDetails As String
Dim HostURL As String
Dim xSheet As Worksheet
Dim headerWS As Worksheet
Dim lineWS As Worksheet
Dim sURL As String
Dim StartRow As Long

Dim Reply, ReplyName, MSG As String
Dim XmlDoc As Object
Dim HttpClient As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

Set headerWS = Sheet1
Set lineWS = Sheet2
Set xSheet = Nothing

If ActiveSheet.CodeName = headerWS.CodeName Then
    Set xSheet = headerWS
Else
    Set xSheet = lineWS
End If

If xSheet.Range("B4").Value = "Production" Then
    HostURL = "https://yourdomain.cloud.infor.com:12345"
Else
    HostURL = "https://yourdomaindev.cloud.infor.com:12345"
End If


xSheet.Range("E6").ClearContents
strUsername = "INFORBC\" & UCase(xSheet.Range("B2").Value)
strPassword = xSheet.Range("B3").Value
StartRow = xSheet.Range("B7")


         USID = CStr(xSheet.Range("B2").Value)
         'If USID <> vbNullString Then
            apiDetails = "?USID=" + UCase(USID)
         'End If

         CONO = CStr(xSheet.Range("E2").Value)
         If CONO <> vbNullString Then
            apiDetails = apiDetails + "&CONO=" + UCase(CONO)
         Else
            MsgBox "Company missing. Please enter a company and try again!", vbInformation, "M3 Error"
            Exit Sub
         End If

         DIVI = CStr(xSheet.Range("E3").Value)
         If DIVI <> vbNullString Then
            apiDetails = apiDetails + "&DIVI=" + UCase(DIVI)
         Else
            xSheet.Range("E3").Value = CStr(xSheet.Range("E" & StartRow).Value)
            apiDetails = apiDetails + "&DIVI=" + CStr(xSheet.Range("E" & StartRow).Value)
         End If

         FACI = CStr(xSheet.Range("E4").Value)
         If FACI <> vbNullString Then
            apiDetails = apiDetails + "&FACI=" + UCase(FACI)
         Else
            xSheet.Range("E4").Value = CStr(xSheet.Range("E" & StartRow).Value)
            apiDetails = apiDetails + "&FACI=" + CStr(xSheet.Range("E" & StartRow).Value)
         End If

         WHLO = CStr(xSheet.Range("E5").Value)
         If WHLO <> vbNullString Then
            apiDetails = apiDetails + "&WHLO=" + UCase(WHLO)
         Else
            xSheet.Range("E5").Value = CStr(xSheet.Range("H" & StartRow).Value)
            apiDetails = apiDetails + "&WHLO=" + CStr(xSheet.Range("H" & StartRow).Value)
         End If

sURL = HostURL + "/m3api-rest/execute/MNS150MI/" & "ChgDefaultValue" & apiDetails

With HttpClient
    .Open "GET", sURL, False, strUsername, strPassword
    .setRequestHeader "Content-Type", "application/xml"
    .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store chache
    .setRequestHeader "Authorization", "Basic " + Encoding.Base64Encode(strUsername + ":" + strPassword)
    .send 'send HTTP request
    Reply = .responseText
End With

    If HttpClient.Status = 200 Then ' Success (200)
        XmlDoc.LoadXML Reply
        ReplyName = XmlDoc.DocumentElement.nodeName
             If ReplyName <> "ErrorMessage" Then

            'MsgBox "OK " & HttpClient.Status
            With XmlDoc
                MSG = .DocumentElement.FirstChild.Text
                'xSheet.Range("E6").Value = "OK"
                xSheet.Range("E6").Value = MSG & " Updated OK"
                xSheet.Range("E6").Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                xSheet.Range("E6").Replace "  ", ""

                If ActiveSheet.CodeName = headerWS.CodeName Then
                    Head.responseError = MSG
                    Head.UpdateUserProfileError = False
                Else
                    Line.responseError = MSG
                    Line.UpdateUserProfileError = False
                End If
                End With
            Else ' Failure (404, 500, ...)
            With XmlDoc
                MSG = .DocumentElement.FirstChild.Text
                'Debug.Print msg
                'xSheet.Range("E6").Value = "NOK"
                xSheet.Range("E6").Value = MSG & " NOK"
                xSheet.Range("E6").Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                xSheet.Range("E6").Replace "  ", ""

                If ActiveSheet.CodeName = headerWS.CodeName Then
                    Head.responseError = MSG
                    Head.UpdateUserProfileError = True
                Else
                    Line.responseError = MSG
                    Line.UpdateUserProfileError = True
                End If
              End With
            End If
    Else
        Application.ScreenUpdating = True
        xSheet.Range("E6").Value = HttpClient.Status & " " & HttpClient.statusText
        MsgBox "Error: " & HttpClient.Status & " " & HttpClient.statusText, vbCritical, "MNS150MI"
    Exit Sub
    End If
End Sub
