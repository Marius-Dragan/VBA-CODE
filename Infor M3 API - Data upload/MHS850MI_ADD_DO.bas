Attribute VB_Name = "MHS850MI_ADD_DO"
Option Explicit
Sub Upload_MHS850MI_AddDO()

Dim M3Response, M3ResponseName, MSG As String
Dim XmlDoc As Object
Dim HttpClient As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

'Configure sheet and tranzaction
Dim strTransaction As String
Dim xSheet As Worksheet
Set xSheet = Sheet1
strTransaction = xSheet.Range("B5")

'Configure settings
Dim strUsername
strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
Dim strPassword
strPassword = xSheet.Range("B3")

'Configure environment
Dim strProgram As String
strProgram = "MHS850MI"
Dim sURL As String
Dim envURL As String
If xSheet.Range("B4").Value = "Production" Then
    envURL = "https://yourdomain.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
Else
    envURL = "https://yourdomaindev.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
End If

'Capture the starting Row
Dim StartRow As Long
StartRow = xSheet.Range("B7")
Dim EndRow As Long
EndRow = xSheet.Range("B8")

Application.ScreenUpdating = False
Do While (StartRow <= EndRow)

    sURL = envURL
    sURL = sURL & "&PRMD=" & xSheet.Cells(9, 2)
    sURL = sURL & "&CONO=" & xSheet.Cells(StartRow, 3)
    sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&ITNO=" & xSheet.Cells(StartRow, 5)
    sURL = sURL & "&WHSL=" & xSheet.Cells(StartRow, 6)

    If xSheet.Cells(StartRow, 7).Value <> vbNullString Then
        sURL = sURL & "&TWSL=" & xSheet.Cells(StartRow, 7)
    End If

    sURL = sURL & "&DLQT=" & xSheet.Cells(StartRow, 8)
    sURL = sURL & "&TRTP=" & xSheet.Cells(StartRow, 9)
    sURL = sURL & "&RESP=" & xSheet.Cells(StartRow, 10)
    sURL = sURL & "&RSCD=" & xSheet.Cells(StartRow, 11)


    If xSheet.Cells(StartRow, 12).Value <> vbNullString Then
        sURL = sURL & "&MSGN=" & xSheet.Cells(StartRow, 12)
    End If

    If xSheet.Cells(StartRow, 13).Value <> vbNullString Then
        sURL = sURL & "&PACN=" & xSheet.Cells(StartRow, 13)
    End If

    If xSheet.Cells(StartRow, 14).Value <> vbNullString Then
        sURL = sURL & "&GEDT=" & xSheet.Cells(StartRow, 14)
    End If

    If xSheet.Cells(StartRow, 15).Value <> vbNullString Then
        sURL = sURL & "&GETM=" & xSheet.Cells(StartRow, 15)
    End If

    If xSheet.Cells(StartRow, 16).Value <> vbNullString Then
        sURL = sURL & "&E0PA=" & xSheet.Cells(StartRow, 16)
    End If

    If xSheet.Cells(StartRow, 17).Value <> vbNullString Then
        sURL = sURL & "&E0PB=" & xSheet.Cells(StartRow, 17)
    End If

    If xSheet.Cells(StartRow, 18).Value <> vbNullString Then
        sURL = sURL & "&E065=" & xSheet.Cells(StartRow, 18)
    End If

    If xSheet.Cells(StartRow, 19).Value <> vbNullString Then
        sURL = sURL & "&CUNO=" & xSheet.Cells(StartRow, 19)
    End If

    If xSheet.Cells(StartRow, 20).Value <> vbNullString Then
        sURL = sURL & "&ADID=" & xSheet.Cells(StartRow, 20)
    End If

    If xSheet.Cells(StartRow, 21).Value <> vbNullString Then
        sURL = sURL & "&POPN=" & xSheet.Cells(StartRow, 21)
    End If

    If xSheet.Cells(StartRow, 22).Value <> vbNullString Then
        sURL = sURL & "&ALWQ=" & xSheet.Cells(StartRow, 22)
    End If

    If xSheet.Cells(StartRow, 23).Value <> vbNullString Then
        sURL = sURL & "&ALWT=" & xSheet.Cells(StartRow, 23)
    End If

    If xSheet.Cells(StartRow, 24).Value <> vbNullString Then
        sURL = sURL & "&BANO=" & xSheet.Cells(StartRow, 24)
    End If

    If xSheet.Cells(StartRow, 25).Value <> vbNullString Then
        sURL = sURL & "&CAMU=" & xSheet.Cells(StartRow, 25)
    End If

    If xSheet.Cells(StartRow, 26).Value <> vbNullString Then
        sURL = sURL & "&ALQT=" & xSheet.Cells(StartRow, 26)
    End If

    If xSheet.Cells(StartRow, 27).Value <> vbNullString Then
        sURL = sURL & "&RIDN=" & xSheet.Cells(StartRow, 27)
    End If

    If xSheet.Cells(StartRow, 28).Value <> vbNullString Then
        sURL = sURL & "&RIDL=" & xSheet.Cells(StartRow, 28)
    End If

    If xSheet.Cells(StartRow, 29).Value <> vbNullString Then
        sURL = sURL & "&RIDX=" & xSheet.Cells(StartRow, 29)
    End If

    If xSheet.Cells(StartRow, 30).Value <> vbNullString Then
        sURL = sURL & "&RIDI=" & xSheet.Cells(StartRow, 30)
    End If

    If xSheet.Cells(StartRow, 31).Value <> vbNullString Then
        sURL = sURL & "&PLSX=" & xSheet.Cells(StartRow, 31)
    End If

    If xSheet.Cells(StartRow, 32).Value <> vbNullString Then
        sURL = sURL & "&DLIX=" & xSheet.Cells(StartRow, 32)
    End If

    If xSheet.Cells(StartRow, 33).Value <> vbNullString Then
        sURL = sURL & "&USD1=" & xSheet.Cells(StartRow, 33)
    End If

    If xSheet.Cells(StartRow, 34).Value <> vbNullString Then
        sURL = sURL & "&USD2=" & xSheet.Cells(StartRow, 34)
    End If

    If xSheet.Cells(StartRow, 35).Value <> vbNullString Then
        sURL = sURL & "&USD3=" & xSheet.Cells(StartRow, 35)
    End If

    If xSheet.Cells(StartRow, 36).Value <> vbNullString Then
        sURL = sURL & "&USD4=" & xSheet.Cells(StartRow, 36)
    End If

    If xSheet.Cells(StartRow, 37).Value <> vbNullString Then
        sURL = sURL & "&USD5=" & xSheet.Cells(StartRow, 37)
    End If

    If xSheet.Cells(StartRow, 38).Value <> vbNullString Then
        sURL = sURL & "&CAWE=" & xSheet.Cells(StartRow, 38)
    End If

    If xSheet.Cells(StartRow, 39).Value <> vbNullString Then
        sURL = sURL & "&PMSN=" & xSheet.Cells(StartRow, 39)
    End If

    If xSheet.Cells(StartRow, 40).Value <> vbNullString Then
        sURL = sURL & "&OPNO=" & xSheet.Cells(StartRow, 40)
    End If

    If xSheet.Cells(StartRow, 41).Value <> vbNullString Then
        sURL = sURL & "&RORC=" & xSheet.Cells(StartRow, 41)
    End If

    If xSheet.Cells(StartRow, 42).Value <> vbNullString Then
        sURL = sURL & "&RORN=" & xSheet.Cells(StartRow, 42)
    End If

    If xSheet.Cells(StartRow, 43).Value <> vbNullString Then
        sURL = sURL & "&RORL=" & xSheet.Cells(StartRow, 43)
    End If

    If xSheet.Cells(StartRow, 44).Value <> vbNullString Then
        sURL = sURL & "&RORX=" & xSheet.Cells(StartRow, 44)
    End If

    If xSheet.Cells(StartRow, 45).Value <> vbNullString Then
        sURL = sURL & "&BREF=" & xSheet.Cells(StartRow, 45)
    End If

    If xSheet.Cells(StartRow, 46).Value <> vbNullString Then
        sURL = sURL & "&BRE2=" & xSheet.Cells(StartRow, 46)
    End If

    If xSheet.Cells(StartRow, 47).Value <> vbNullString Then
        sURL = sURL & "&RPDT=" & xSheet.Cells(StartRow, 47)
    End If

    If xSheet.Cells(StartRow, 48).Value <> vbNullString Then
        sURL = sURL & "&RPTM=" & xSheet.Cells(StartRow, 48)
    End If

    If xSheet.Cells(StartRow, 49).Value <> vbNullString Then
        sURL = sURL & "&REPN=" & xSheet.Cells(StartRow, 49)
    End If

    If xSheet.Cells(StartRow, 50).Value <> vbNullString Then
        sURL = sURL & "&UTCM=" & xSheet.Cells(StartRow, 50)
    End If


    With HttpClient
        .Open "GET", sURL, False, strUsername, strPassword
        .setRequestHeader "Content-Type", "application/xml"
        .setRequestHeader "Cache-Control", "no-cache"
        .setRequestHeader "Authorization", "Basic " + Encoding.Base64Encode(strUsername + ":" + strPassword)
        .send 'send HTTP request
        M3Response = .responseText
    End With

         If HttpClient.Status = 200 Then ' Success (200)
            XmlDoc.LoadXML M3Response
            M3ResponseName = XmlDoc.DocumentElement.nodeName
                 If M3ResponseName <> "ErrorMessage" Then

                'MsgBox "OK " & M3Service.Status
                With XmlDoc
                    MSG = .DocumentElement.FirstChild.Text
                    xSheet.Range("A" & StartRow).Value = "OK"
                    'xSheet.Range("B" & StartRow).Value = MSG & " Uploaded OK"
                    xSheet.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                    xSheet.Range("B" & StartRow).Replace "  ", ""
                    End With
                Else ' Failure (404, 500, ...)
                With XmlDoc
                        MSG = .DocumentElement.FirstChild.Text
                      'Debug.Print msg
                      xSheet.Range("A" & StartRow).Value = "NOK"
                      xSheet.Range("B" & StartRow).Value = MSG
                      xSheet.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                      xSheet.Range("B" & StartRow).Replace "  ", ""
                  End With
                End If
            Else
                MsgBox "Error: " & HttpClient.Status & " " & HttpClient.statusText, vbCritical, strProgram & " " & strTransaction
                Exit Sub
        End If

    sURL = ""
    StartRow = StartRow + 1

Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, strProgram & " " & strTransaction
End Sub
Sub ClearLogs()
Dim LRow As Long
    With Sheet1
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
