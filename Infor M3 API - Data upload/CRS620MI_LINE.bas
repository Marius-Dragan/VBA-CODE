Attribute VB_Name = "CRS620MI_LINE"
Option Explicit
Sub Upload_CRS620_AddSupplier()

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
strProgram = "CRS620MI"
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
    'check number of line matches with front end
    xColumn = 3
    sURL = sURL & "&SUNO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&SUNM=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&SUTY=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ALSU=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CSCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ECAR=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&LNCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DTFM=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&MEPF=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&HAFE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&QUCL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ORTY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEDL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&MODL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEAF=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEPA=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DT4T=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DTCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&VTCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TXAP=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TAXC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CUCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CRTP=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TEPY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ATPR=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ACRF=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PHNO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PHN2=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TFNO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TLNO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CORG=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&COR2=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&VRNO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUCO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DESV=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&FWSC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUAL=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&EDIT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUCM=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PODA=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&BUYE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RESP=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&AGNT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ABSK=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&ABSM=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PWMT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DCSM=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&FUSC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SPFC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&COBI=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SCNO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUGR=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SHST=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&POOT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&OUCN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TINO=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PRSU=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SERS=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SBPE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PACD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PTDY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUST=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DTDY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TECD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&REGR=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SUSY=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SHAC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&AVCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TAME=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&TDCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&IAPT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&IAPC=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&IAPE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&IAPF=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI1=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI2=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI3=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI4=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI5=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI6=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI7=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI8=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CFI9=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CF10=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&STAT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PPIN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CGRP=" & xSheet.Cells(StartRow, xColumn)
    End If

        xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RASN=" & xSheet.Cells(StartRow, xColumn)
    End If

        xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&GEOC=" & xSheet.Cells(StartRow, xColumn)
    End If

        xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&PPLV=" & xSheet.Cells(StartRow, xColumn)
    End If

        xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&CINP=" & xSheet.Cells(StartRow, xColumn)
    End If

        xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&SCIS=" & xSheet.Cells(StartRow, xColumn)
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
    With Sheet2
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
