Attribute VB_Name = "MNS150MI_Change"
Option Explicit
Sub UploadLineMNS150MIChange()

Dim M3Response, M3ResponseName, MSG As String
Dim XmlDoc As Object
Dim HttpClient As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

'Configure sheet and tranzaction
Dim strTransaction As String
Dim xSheet As Worksheet
Set xSheet = Sheet2
strTransaction = xSheet.Range("B5")

'Configure settings
Dim strUsername
strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
Dim strPassword
strPassword = xSheet.Range("B3")

'Configure env
Dim strProgram As String
strProgram = "MNS150MI"
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
    sURL = sURL & "&USID=" & xSheet.Cells(StartRow, 3)
    sURL = sURL & "&CONO=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&DIVI=" & xSheet.Cells(StartRow, 5)

    If xSheet.Cells(StartRow, 6).Value <> "" Then
        sURL = sURL & "&LANC=" & xSheet.Cells(StartRow, 6)
    End If

    If xSheet.Cells(StartRow, 7).Value <> "" Then
        sURL = sURL & "&DTFM=" & xSheet.Cells(StartRow, 7)
    End If

    If xSheet.Cells(StartRow, 8).Value <> "" Then
        sURL = sURL & "&DCFM=" & xSheet.Cells(StartRow, 8)
    End If

    If xSheet.Cells(StartRow, 9).Value <> "" Then
        sURL = sURL & "&TIZO=" & xSheet.Cells(StartRow, 9)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&FACI=" & xSheet.Cells(StartRow, 10)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, 11)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&CUNO=" & xSheet.Cells(StartRow, 12)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&DEPT=" & xSheet.Cells(StartRow, 13)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&NAME=" & xSheet.Cells(StartRow, 14)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&EQAL=" & xSheet.Cells(StartRow, 15)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&USTA=" & xSheet.Cells(StartRow, 16)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
        sURL = sURL & "&EUID=" & xSheet.Cells(StartRow, 17)
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
                'MsgBox "OK " & HttpClient.Status
                With XmlDoc
                    MSG = .DocumentElement.FirstChild.Text
                    xSheet.Range("A" & StartRow).Value = "OK"
                    'xSheet.Range("B" & StartRow).Value = MSG & " Updated OK"
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
Sub ClearLogsChange()
Dim LRow As Long
    With Sheet2
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
