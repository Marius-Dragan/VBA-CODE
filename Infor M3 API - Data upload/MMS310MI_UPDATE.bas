Attribute VB_Name = "MMS310MI_UPDATE"
Option Explicit
Sub UpdateMMS310MI()

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
strProgram = "MMS310MI"
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
    sURL = sURL & "&CONO=" & xSheet.Cells(StartRow, 3)
    sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&ITNO=" & xSheet.Cells(StartRow, 5)

    If xSheet.Cells(StartRow, 6).Value <> vbNullString Then
    sURL = sURL & "&WHSL=" & xSheet.Cells(StartRow, 6)
    End If

    If xSheet.Cells(StartRow, 7).Value <> vbNullString Then
    sURL = sURL & "&BANO=" & xSheet.Cells(StartRow, 7)
    End If

    If xSheet.Cells(StartRow, 8).Value <> vbNullString Then
    sURL = sURL & "&CAMU=" & xSheet.Cells(StartRow, 8)
    End If

    If xSheet.Cells(StartRow, 9).Value <> vbNullString Then
    sURL = sURL & "&REPN=" & xSheet.Cells(StartRow, 9)
    End If

    If xSheet.Cells(StartRow, 10).Value <> vbNullString Then
    sURL = sURL & "&STQI=" & xSheet.Cells(StartRow, 10)
    End If

    If xSheet.Cells(StartRow, 11).Value <> vbNullString Then
    sURL = sURL & "&STAG=" & xSheet.Cells(StartRow, 11)
    End If

    If xSheet.Cells(StartRow, 12).Value <> vbNullString Then
    sURL = sURL & "&CAWI=" & xSheet.Cells(StartRow, 12)
    End If

    If xSheet.Cells(StartRow, 13).Value <> vbNullString Then
    sURL = sURL & "&STDI=" & xSheet.Cells(StartRow, 13)
    End If

    If xSheet.Cells(StartRow, 14).Value <> vbNullString Then
    sURL = sURL & "&TIHH=" & xSheet.Cells(StartRow, 14)
    End If

    If xSheet.Cells(StartRow, 15).Value <> vbNullString Then
    sURL = sURL & "&TIMM=" & xSheet.Cells(StartRow, 15)
    End If

    If xSheet.Cells(StartRow, 16).Value <> vbNullString Then
    sURL = sURL & "&TISS=" & xSheet.Cells(StartRow, 16)
    End If

    If xSheet.Cells(StartRow, 17).Value <> vbNullString Then
    sURL = sURL & "&PRDT=" & xSheet.Cells(StartRow, 17)
    End If

    If xSheet.Cells(StartRow, 18).Value <> vbNullString Then
    sURL = sURL & "&TRPR=" & xSheet.Cells(StartRow, 18)
    End If

    If xSheet.Cells(StartRow, 19).Value <> vbNullString Then
    sURL = sURL & "&BREF=" & xSheet.Cells(StartRow, 19)
    End If

    If xSheet.Cells(StartRow, 20).Value <> vbNullString Then
    sURL = sURL & "&BRE2=" & xSheet.Cells(StartRow, 20)
    End If

    If xSheet.Cells(StartRow, 21).Value <> vbNullString Then
    sURL = sURL & "&BREM=" & xSheet.Cells(StartRow, 21)
    End If

    If xSheet.Cells(StartRow, 22).Value <> vbNullString Then
    sURL = sURL & "&RSCD=" & xSheet.Cells(StartRow, 22)
    End If


    With HttpClient
        .Open "GET", sURL, False, strUsername, strPassword
        .setRequestHeader "Content-Type", "application/xml"
        .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store cache
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
