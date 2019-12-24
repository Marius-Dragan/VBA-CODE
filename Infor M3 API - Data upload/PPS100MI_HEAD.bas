Attribute VB_Name = "PPS100MI_HEAD"
Option Explicit
Sub Upload_PPS100MI_Head()

Dim M3Response, M3ResponseName, MSG As String
Dim XmlDoc As Object
Dim HttpClient As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

'Configure sheet and tranzaction
Dim xSheet As Worksheet
Dim strTransaction As String
Set xSheet = Sheet1
strTransaction = xSheet.Range("B5")

'Configure settings
Dim strUsername
strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
Dim strPassword
strPassword = xSheet.Range("B3")

'Configure env
Dim strProgram As String
strProgram = "PPS100MI"
Dim sURL As String
Dim envURL As String
If xSheet.Range("B4").Value = "Production" Then
    envURL = "https://yourdomain.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
Else
    envURL = "https://yourdomaindev.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
End If

'Capture the starting Row and current Column
Dim StartRow As Long
StartRow = xSheet.Range("B7")
Dim EndRow As Long
EndRow = xSheet.Range("B8")
Dim xColumn As Long

Application.ScreenUpdating = False
Do While (StartRow <= EndRow)

    sURL = envURL

    xColumn = 3
    sURL = sURL & "&CONO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&SUNO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&AGTP=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&FVDT=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&UVDT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&AGNB=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TX30=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&AGRD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RNDT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PAST=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TENT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&BUYE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&AGPT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RFID=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&QREM=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CUCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEPA=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CRTP=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&MODL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEDL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEPY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEAF=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CIVC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&VAGN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DIP2=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&FACI=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SBAN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ACGR=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PGDP=" & xSheet.Cells(StartRow, xColumn)
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
                    'xSheet.Range("B" & StartRow).NumberFormat = "@"
                    'xSheet.Range("B" & StartRow).Value = Right(.Text, 10)
                    'xSheet.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                    'xSheet.Range("B" & StartRow).Replace "  ", ""
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
Sub ClearLogsHeader()
Dim LRow As Long
    With Sheet1
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
