Attribute VB_Name = "OIS390MI_ADD_HEAD"
Option Explicit
Sub Upload_OIS390MI_AddHead()

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
strProgram = "OIS390MI"
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
    sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&CUNO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&REPN=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&FACI=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&PYNO=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    sURL = sURL & "&ORTP=" & xSheet.Cells(StartRow, xColumn)

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RORN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&EXIN=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&YEA4=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&DIVI=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&EPDT=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RPA1=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RPA2=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RSCD=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&RERE=" & xSheet.Cells(StartRow, xColumn)
    End If

    xColumn = xColumn + 1
    If xSheet.Cells(StartRow, xColumn).Value <> "" Then
        sURL = sURL & "&YREF=" & xSheet.Cells(StartRow, xColumn)
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
