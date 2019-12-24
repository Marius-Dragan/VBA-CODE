Attribute VB_Name = "OIS100MI_LINE"
Option Explicit
Sub UploadLineOIS100Line()

Dim M3Response, M3ResponseName, MSG As String
Dim XmlDoc As Object
Dim HttpClient As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

'Configure sheet and tranzaction
Dim xSheet As Worksheet
Dim strTransaction As String
Set xSheet = Sheet2
strTransaction = xSheet.Range("B5")

'Configure settings
Dim strUsername
strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
Dim strPassword
strPassword = xSheet.Range("B3")

'Configure env
Dim strProgram As String
strProgram = "OIS100MI"
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
    sURL = sURL & "&ORNO=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&ITNO=" & xSheet.Cells(StartRow, 5)
    sURL = sURL & "&ORQT=" & xSheet.Cells(StartRow, 6)
    sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, 7)
    sURL = sURL & "&DWDT=" & xSheet.Cells(StartRow, 8)


    If xSheet.Cells(StartRow, 9).Value <> "" Then
        sURL = sURL & "&JDCD=" & xSheet.Cells(StartRow, 9)
    End If

    sURL = sURL & "&CUPO=" & xSheet.Cells(StartRow, 10)
    sURL = sURL & "&SAPR=" & xSheet.Cells(StartRow, 11)

    If xSheet.Cells(StartRow, 12).Value <> "" Then
        sURL = sURL & "&DIA1=" & xSheet.Cells(StartRow, 12)
    End If

    If xSheet.Cells(StartRow, 13).Value <> "" Then
        sURL = sURL & "&DIA2=" & xSheet.Cells(StartRow, 13)
    End If

    If xSheet.Cells(StartRow, 14).Value <> "" Then
        sURL = sURL & "&DIA3=" & xSheet.Cells(StartRow, 14)
    End If

    If xSheet.Cells(StartRow, 15).Value <> "" Then
        sURL = sURL & "&DIA4=" & xSheet.Cells(StartRow, 15)
    End If

    If xSheet.Cells(StartRow, 16).Value <> "" Then
        sURL = sURL & "&DIA5=" & xSheet.Cells(StartRow, 16)
    End If

    If xSheet.Cells(StartRow, 17).Value <> "" Then
        sURL = sURL & "&DIA6=" & xSheet.Cells(StartRow, 17)
    End If

    If xSheet.Cells(StartRow, 18).Value <> "" Then
        sURL = sURL & "&DLSP=" & xSheet.Cells(StartRow, 18)
    End If

    If xSheet.Cells(StartRow, 19).Value <> "" Then
        sURL = sURL & "&DLSX=" & xSheet.Cells(StartRow, 19)
    End If

    If xSheet.Cells(StartRow, 20).Value <> "" Then
        sURL = sURL & "&CFXX=" & xSheet.Cells(StartRow, 20)
    End If

    If xSheet.Cells(StartRow, 21).Value <> "" Then
        sURL = sURL & "&ECVS=" & xSheet.Cells(StartRow, 21)
    End If

    If xSheet.Cells(StartRow, 22).Value <> "" Then
        sURL = sURL & "&ALUN=" & xSheet.Cells(StartRow, 22)
    End If

    If xSheet.Cells(StartRow, 23).Value <> "" Then
        sURL = sURL & "&CODT=" & xSheet.Cells(StartRow, 23)
    End If

    If xSheet.Cells(StartRow, 24).Value <> "" Then
        sURL = sURL & "&ITDS=" & xSheet.Cells(StartRow, 24)
    End If

    If xSheet.Cells(StartRow, 25).Value <> "" Then
        sURL = sURL & "&DIP1=" & xSheet.Cells(StartRow, 25)
    End If

    If xSheet.Cells(StartRow, 26).Value <> "" Then
        sURL = sURL & "&DIP2=" & xSheet.Cells(StartRow, 26)
    End If

    If xSheet.Cells(StartRow, 27).Value <> "" Then
        sURL = sURL & "&DIP3=" & xSheet.Cells(StartRow, 27)
    End If

    If xSheet.Cells(StartRow, 28).Value <> "" Then
        sURL = sURL & "&DIP4=" & xSheet.Cells(StartRow, 28)
    End If

    If xSheet.Cells(StartRow, 29).Value <> "" Then
        sURL = sURL & "&DIP5=" & xSheet.Cells(StartRow, 29)
    End If

    If xSheet.Cells(StartRow, 30).Value <> "" Then
        sURL = sURL & "&DIP6=" & xSheet.Cells(StartRow, 30)
    End If

    If xSheet.Cells(StartRow, 31).Value <> "" Then
        sURL = sURL & "&ALWT=" & xSheet.Cells(StartRow, 31)
    End If

    If xSheet.Cells(StartRow, 32).Value <> "" Then
        sURL = sURL & "&ALWQ=" & xSheet.Cells(StartRow, 32)
    End If

    If xSheet.Cells(StartRow, 33).Value <> "" Then
        sURL = sURL & "&AGNO=" & xSheet.Cells(StartRow, 33)
    End If

    If xSheet.Cells(StartRow, 34).Value <> "" Then
        sURL = sURL & "&CAMU=" & xSheet.Cells(StartRow, 34)
    End If

    If xSheet.Cells(StartRow, 35).Value <> "" Then
        sURL = sURL & "&PROJ=" & xSheet.Cells(StartRow, 35)
    End If

    If xSheet.Cells(StartRow, 36).Value <> "" Then
        sURL = sURL & "&ELNO=" & xSheet.Cells(StartRow, 36)
    End If

    If xSheet.Cells(StartRow, 37).Value <> "" Then
        sURL = sURL & "&CUOR=" & xSheet.Cells(StartRow, 37)
    End If

    If xSheet.Cells(StartRow, 38).Value <> "" Then
        sURL = sURL & "&CUPA=" & xSheet.Cells(StartRow, 38)
    End If

    If xSheet.Cells(StartRow, 39).Value <> "" Then
        sURL = sURL & "&DWHM=" & xSheet.Cells(StartRow, 39)
    End If

    If xSheet.Cells(StartRow, 40).Value <> "" Then
        sURL = sURL & "&D1QT=" & xSheet.Cells(StartRow, 40)
    End If

    If xSheet.Cells(StartRow, 41).Value <> "" Then
        sURL = sURL & "&PACT=" & xSheet.Cells(StartRow, 41)
    End If

    If xSheet.Cells(StartRow, 42).Value <> "" Then
        sURL = sURL & "&POPN=" & xSheet.Cells(StartRow, 42)
    End If

    If xSheet.Cells(StartRow, 43).Value <> "" Then
        sURL = sURL & "&SACD=" & xSheet.Cells(StartRow, 43)
    End If

    If xSheet.Cells(StartRow, 44).Value <> "" Then
        sURL = sURL & "&SPUN=" & xSheet.Cells(StartRow, 44)
    End If

    If xSheet.Cells(StartRow, 45).Value <> "" Then
        sURL = sURL & "&TEPA=" & xSheet.Cells(StartRow, 45)
    End If

    If xSheet.Cells(StartRow, 46).Value <> "" Then
        sURL = sURL & "&EDFP=" & xSheet.Cells(StartRow, 46)
    End If

    If xSheet.Cells(StartRow, 47).Value <> "" Then
        sURL = sURL & "&DWDZ=" & xSheet.Cells(StartRow, 47)
    End If

    If xSheet.Cells(StartRow, 48).Value <> "" Then
        sURL = sURL & "&DWHZ=" & xSheet.Cells(StartRow, 48)
    End If

    If xSheet.Cells(StartRow, 49).Value <> "" Then
        sURL = sURL & "&COHM=" & xSheet.Cells(StartRow, 49)
    End If

    If xSheet.Cells(StartRow, 50).Value <> "" Then
        sURL = sURL & "&CODZ=" & xSheet.Cells(StartRow, 50)
    End If

    If xSheet.Cells(StartRow, 51).Value <> "" Then
        sURL = sURL & "&COHZ=" & xSheet.Cells(StartRow, 51)
    End If

    If xSheet.Cells(StartRow, 52).Value <> "" Then
        sURL = sURL & "&HDPR=" & xSheet.Cells(StartRow, 52)
    End If

    If xSheet.Cells(StartRow, 53).Value <> "" Then
        sURL = sURL & "&ADID=" & xSheet.Cells(StartRow, 53)
    End If

    If xSheet.Cells(StartRow, 54).Value <> "" Then
        sURL = sURL & "&CUSX=" & xSheet.Cells(StartRow, 54)
    End If

    If xSheet.Cells(StartRow, 55).Value <> "" Then
        sURL = sURL & "&DIC1=" & xSheet.Cells(StartRow, 55)
    End If

    If xSheet.Cells(StartRow, 56).Value <> "" Then
        sURL = sURL & "&DIC2=" & xSheet.Cells(StartRow, 56)
    End If

    If xSheet.Cells(StartRow, 57).Value <> "" Then
        sURL = sURL & "&DIC3=" & xSheet.Cells(StartRow, 57)
    End If

    If xSheet.Cells(StartRow, 58).Value <> "" Then
        sURL = sURL & "&DIC4=" & xSheet.Cells(StartRow, 58)
    End If

    If xSheet.Cells(StartRow, 59).Value <> "" Then
        sURL = sURL & "&DIC5=" & xSheet.Cells(StartRow, 59)
    End If

    If xSheet.Cells(StartRow, 60).Value <> "" Then
        sURL = sURL & "&DIC6=" & xSheet.Cells(StartRow, 60)
    End If

    If xSheet.Cells(StartRow, 61).Value <> "" Then
        sURL = sURL & "&CMNO=" & xSheet.Cells(StartRow, 61)
    End If

    If xSheet.Cells(StartRow, 62).Value <> "" Then
        sURL = sURL & "&RSCD=" & xSheet.Cells(StartRow, 62)
    End If

    If xSheet.Cells(StartRow, 63).Value <> "" Then
        sURL = sURL & "&TEDS=" & xSheet.Cells(StartRow, 63)
    End If

    If xSheet.Cells(StartRow, 64).Value <> "" Then
        sURL = sURL & "&CFIN=" & xSheet.Cells(StartRow, 64)
    End If

    If xSheet.Cells(StartRow, 65).Value <> "" Then
        sURL = sURL & "&BANO=" & xSheet.Cells(StartRow, 65)
    End If

    If xSheet.Cells(StartRow, 66).Value <> "" Then
        sURL = sURL & "&WHSL=" & xSheet.Cells(StartRow, 66)
    End If

    If xSheet.Cells(StartRow, 67).Value <> "" Then
        sURL = sURL & "&PRHL=" & xSheet.Cells(StartRow, 67)
    End If

    If xSheet.Cells(StartRow, 68).Value <> "" Then
        sURL = sURL & "&SERN=" & xSheet.Cells(StartRow, 68)
    End If

    If xSheet.Cells(StartRow, 69).Value <> "" Then
        sURL = sURL & "&CTNO=" & xSheet.Cells(StartRow, 69)
    End If

    If xSheet.Cells(StartRow, 70).Value <> "" Then
        sURL = sURL & "&CFGL=" & xSheet.Cells(StartRow, 70)
    End If

    If xSheet.Cells(StartRow, 71).Value <> "" Then
        sURL = sURL & "&GWTP=" & xSheet.Cells(StartRow, 71)
    End If

    If xSheet.Cells(StartRow, 72).Value <> "" Then
        sURL = sURL & "&WATP=" & xSheet.Cells(StartRow, 72)
    End If

    If xSheet.Cells(StartRow, 73).Value <> "" Then
        sURL = sURL & "&PRHW=" & xSheet.Cells(StartRow, 73)
    End If

    If xSheet.Cells(StartRow, 74).Value <> "" Then
        sURL = sURL & "&SERW=" & xSheet.Cells(StartRow, 74)
    End If

    If xSheet.Cells(StartRow, 75).Value <> "" Then
        sURL = sURL & "&PWNR=" & xSheet.Cells(StartRow, 75)
    End If

    If xSheet.Cells(StartRow, 76).Value <> "" Then
        sURL = sURL & "&PWSX=" & xSheet.Cells(StartRow, 76)
    End If

    If xSheet.Cells(StartRow, 77).Value <> "" Then
        sURL = sURL & "&EWST=" & xSheet.Cells(StartRow, 77)
    End If

    If xSheet.Cells(StartRow, 78).Value <> "" Then
        sURL = sURL & "&DANR=" & xSheet.Cells(StartRow, 78)
    End If

    If xSheet.Cells(StartRow, 79).Value <> "" Then
        sURL = sURL & "&TECN=" & xSheet.Cells(StartRow, 79)
    End If

    If xSheet.Cells(StartRow, 80).Value <> "" Then
        sURL = sURL & "&INAP=" & xSheet.Cells(StartRow, 80)
    End If

    If xSheet.Cells(StartRow, 81).Value <> "" Then
        sURL = sURL & "&DRDN=" & xSheet.Cells(StartRow, 81)
    End If

    If xSheet.Cells(StartRow, 82).Value <> "" Then
        sURL = sURL & "&DRDL=" & xSheet.Cells(StartRow, 82)
    End If

    If xSheet.Cells(StartRow, 83).Value <> "" Then
        sURL = sURL & "&DRDX=" & xSheet.Cells(StartRow, 83)
    End If

    If xSheet.Cells(StartRow, 84).Value <> "" Then
        sURL = sURL & "&PIDE=" & xSheet.Cells(StartRow, 84)
    End If

    If xSheet.Cells(StartRow, 85).Value <> "" Then
        sURL = sURL & "&TEPY=" & xSheet.Cells(StartRow, 85)
    End If

    If xSheet.Cells(StartRow, 86).Value <> "" Then
        sURL = sURL & "&LTYP=" & xSheet.Cells(StartRow, 86)
    End If

    If xSheet.Cells(StartRow, 87).Value <> "" Then
        sURL = sURL & "&EXH2=" & xSheet.Cells(StartRow, 87)
    End If

    If xSheet.Cells(StartRow, 88).Value <> "" Then
        sURL = sURL & "&EXC2=" & xSheet.Cells(StartRow, 88)
    End If

    If xSheet.Cells(StartRow, 89).Value <> "" Then
        sURL = sURL & "&MCHP=" & xSheet.Cells(StartRow, 89)
    End If

    If xSheet.Cells(StartRow, 90).Value <> "" Then
        sURL = sURL & "&EXHE=" & xSheet.Cells(StartRow, 90)
    End If

    If xSheet.Cells(StartRow, 91).Value <> "" Then
        sURL = sURL & "&EXCN=" & xSheet.Cells(StartRow, 91)
    End If

    If xSheet.Cells(StartRow, 92).Value <> "" Then
        sURL = sURL & "&MODL=" & xSheet.Cells(StartRow, 92)
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
Sub ClearLogsLines()
Dim LRow As Long
    With Sheet2
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
