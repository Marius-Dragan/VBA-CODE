Attribute VB_Name = "OIS100MI_HEAD"
Option Explicit
Sub UploadLineOIS100Head()

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
    sURL = sURL & "&CUNO=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&ORTP=" & xSheet.Cells(StartRow, 5)
    sURL = sURL & "&FACI=" & xSheet.Cells(StartRow, 6)
    sURL = sURL & "&PROJ=" & xSheet.Cells(StartRow, 7)
    sURL = sURL & "&ELNO=" & xSheet.Cells(StartRow, 8)
    sURL = sURL & "&WHLO=" & xSheet.Cells(StartRow, 9)
    sURL = sURL & "&RESP=" & xSheet.Cells(StartRow, 10)
    sURL = sURL & "&LNCD=" & xSheet.Cells(StartRow, 11)
    sURL = sURL & "&RLDT=" & xSheet.Cells(StartRow, 12)
    sURL = sURL & "&CUOR=" & xSheet.Cells(StartRow, 13)
    sURL = sURL & "&MODL=" & xSheet.Cells(StartRow, 14)
    sURL = sURL & "&TEDL=" & xSheet.Cells(StartRow, 15)
    sURL = sURL & "&TEPY=" & xSheet.Cells(StartRow, 16)
    sURL = sURL & "&CUCD=" & xSheet.Cells(StartRow, 17)
    sURL = sURL & "&ORDT=" & xSheet.Cells(StartRow, 18)
    sURL = sURL & "&TEPA=" & xSheet.Cells(StartRow, 19)
    sURL = sURL & "&RLDZ=" & xSheet.Cells(StartRow, 20)
    sURL = sURL & "&FDED=" & xSheet.Cells(StartRow, 21)
    sURL = sURL & "&LDED=" & xSheet.Cells(StartRow, 22)

    If xSheet.Cells(StartRow, 23).Value <> "" Then
        sURL = sURL & "&AGNT=" & xSheet.Cells(StartRow, 23)
    End If

    If xSheet.Cells(StartRow, 24).Value <> "" Then
        sURL = sURL & "&SMCD=" & xSheet.Cells(StartRow, 24)
    End If

    If xSheet.Cells(StartRow, 25).Value <> "" Then
        sURL = sURL & "&YREF=" & xSheet.Cells(StartRow, 25)
    End If

    If xSheet.Cells(StartRow, 26).Value <> "" Then
        sURL = sURL & "&OTBA=" & xSheet.Cells(StartRow, 26)
    End If

    If xSheet.Cells(StartRow, 27).Value <> "" Then
        sURL = sURL & "&PYNO=" & xSheet.Cells(StartRow, 27)
    End If

    If xSheet.Cells(StartRow, 28).Value <> "" Then
        sURL = sURL & "&ADID=" & xSheet.Cells(StartRow, 28)
    End If

    If xSheet.Cells(StartRow, 29).Value <> "" Then
        sURL = sURL & "&OREF=" & xSheet.Cells(StartRow, 29)
    End If

    If xSheet.Cells(StartRow, 30).Value <> "" Then
        sURL = sURL & "&OFNO=" & xSheet.Cells(StartRow, 30)
    End If

    If xSheet.Cells(StartRow, 31).Value <> "" Then
        sURL = sURL & "&TEL2=" & xSheet.Cells(StartRow, 31)
    End If

    If xSheet.Cells(StartRow, 32).Value <> "" Then
        sURL = sURL & "&CUDT=" & xSheet.Cells(StartRow, 32)
    End If

    If xSheet.Cells(StartRow, 33).Value <> "" Then
        sURL = sURL & "&INRC=" & xSheet.Cells(StartRow, 33)
    End If

    If xSheet.Cells(StartRow, 34).Value <> "" Then
        sURL = sURL & "&RLHM=" & xSheet.Cells(StartRow, 34)
    End If

    If xSheet.Cells(StartRow, 35).Value <> "" Then
        sURL = sURL & "&TIZO=" & xSheet.Cells(StartRow, 35)
    End If

    If xSheet.Cells(StartRow, 36).Value <> "" Then
        sURL = sURL & "&AGNO=" & xSheet.Cells(StartRow, 36)
    End If

    If xSheet.Cells(StartRow, 37).Value <> "" Then
        sURL = sURL & "&PRO2=" & xSheet.Cells(StartRow, 37)
    End If

    If xSheet.Cells(StartRow, 38).Value <> "" Then
        sURL = sURL & "&RLHZ=" & xSheet.Cells(StartRow, 38)
    End If

    If xSheet.Cells(StartRow, 39).Value <> "" Then
        sURL = sURL & "&DLSP=" & xSheet.Cells(StartRow, 39)
    End If

    If xSheet.Cells(StartRow, 40).Value <> "" Then
        sURL = sURL & "&DSTX=" & xSheet.Cells(StartRow, 40)
    End If

    If xSheet.Cells(StartRow, 41).Value <> "" Then
        sURL = sURL & "&TECD=" & xSheet.Cells(StartRow, 41)
    End If

    If xSheet.Cells(StartRow, 42).Value <> "" Then
        sURL = sURL & "&PLTB=" & xSheet.Cells(StartRow, 42)
    End If

    If xSheet.Cells(StartRow, 43).Value <> "" Then
        sURL = sURL & "&DISY=" & xSheet.Cells(StartRow, 43)
    End If

    If xSheet.Cells(StartRow, 44).Value <> "" Then
        sURL = sURL & "&WCON=" & xSheet.Cells(StartRow, 44)
    End If

    If xSheet.Cells(StartRow, 45).Value <> "" Then
        sURL = sURL & "&ID01=" & xSheet.Cells(StartRow, 45)
    End If

    If xSheet.Cells(StartRow, 46).Value <> "" Then
        sURL = sURL & "&ID02=" & xSheet.Cells(StartRow, 46)
    End If

    If xSheet.Cells(StartRow, 47).Value <> "" Then
        sURL = sURL & "&OTDP=" & xSheet.Cells(StartRow, 47)
    End If

    If xSheet.Cells(StartRow, 48).Value <> "" Then
        sURL = sURL & "&CRTP=" & xSheet.Cells(StartRow, 48)
    End If

    If xSheet.Cells(StartRow, 49).Value <> "" Then
        sURL = sURL & "&DICD=" & xSheet.Cells(StartRow, 49)
    End If

    If xSheet.Cells(StartRow, 50).Value <> "" Then
        sURL = sURL & "&CHL1=" & xSheet.Cells(StartRow, 50)
    End If

    If xSheet.Cells(StartRow, 51).Value <> "" Then
        sURL = sURL & "&CHL2=" & xSheet.Cells(StartRow, 51)
    End If

    If xSheet.Cells(StartRow, 52).Value <> "" Then
        sURL = sURL & "&CHL3=" & xSheet.Cells(StartRow, 52)
    End If

    If xSheet.Cells(StartRow, 53).Value <> "" Then
        sURL = sURL & "&CHL4=" & xSheet.Cells(StartRow, 53)
    End If

    If xSheet.Cells(StartRow, 54).Value <> "" Then
        sURL = sURL & "&NREF=" & xSheet.Cells(StartRow, 54)
    End If

    If xSheet.Cells(StartRow, 55).Value <> "" Then
        sURL = sURL & "&TRDP=" & xSheet.Cells(StartRow, 55)
    End If

    If xSheet.Cells(StartRow, 56).Value <> "" Then
        sURL = sURL & "&IPAD=" & xSheet.Cells(StartRow, 56)
    End If

    If xSheet.Cells(StartRow, 57).Value <> "" Then
        sURL = sURL & "&ESOV=" & xSheet.Cells(StartRow, 57)
    End If

    If xSheet.Cells(StartRow, 58).Value <> "" Then
        sURL = sURL & "&EXCD=" & xSheet.Cells(StartRow, 58)
    End If

    If xSheet.Cells(StartRow, 59).Value <> "" Then
        sURL = sURL & "&PYCD=" & xSheet.Cells(StartRow, 59)
    End If

    If xSheet.Cells(StartRow, 60).Value <> "" Then
        sURL = sURL & "&FRE1=" & xSheet.Cells(StartRow, 60)
    End If

    If xSheet.Cells(StartRow, 61).Value <> "" Then
        sURL = sURL & "&BREC=" & xSheet.Cells(StartRow, 61)
    End If

    If xSheet.Cells(StartRow, 62).Value <> "" Then
        sURL = sURL & "&OHEA=" & xSheet.Cells(StartRow, 62)
    End If

    If xSheet.Cells(StartRow, 63).Value <> "" Then
        sURL = sURL & "&CUCH=" & xSheet.Cells(StartRow, 63)
    End If

    If xSheet.Cells(StartRow, 64).Value <> "" Then
        sURL = sURL & "&CCAC=" & xSheet.Cells(StartRow, 64)
    End If

    If xSheet.Cells(StartRow, 65).Value <> "" Then
        sURL = sURL & "&DECU=" & xSheet.Cells(StartRow, 65)
    End If

    If xSheet.Cells(StartRow, 66).Value <> "" Then
        sURL = sURL & "&GCAC=" & xSheet.Cells(StartRow, 66)
    End If

    If xSheet.Cells(StartRow, 67).Value <> "" Then
        sURL = sURL & "&PYRE=" & xSheet.Cells(StartRow, 67)
    End If

    If xSheet.Cells(StartRow, 68).Value <> "" Then
        sURL = sURL & "&BKID=" & xSheet.Cells(StartRow, 68)
    End If

    If xSheet.Cells(StartRow, 69).Value <> "" Then
        sURL = sURL & "&OPRI=" & xSheet.Cells(StartRow, 69)
    End If

    If xSheet.Cells(StartRow, 70).Value <> "" Then
        sURL = sURL & "&SPLM=" & xSheet.Cells(StartRow, 70)
    End If

    If xSheet.Cells(StartRow, 71).Value <> "" Then
        sURL = sURL & "&CHSY=" & xSheet.Cells(StartRow, 71)
    End If

    If xSheet.Cells(StartRow, 72).Value <> "" Then
        sURL = sURL & "&OIVR=" & xSheet.Cells(StartRow, 72)
    End If

    If xSheet.Cells(StartRow, 73).Value <> "" Then
        sURL = sURL & "&OYEA=" & xSheet.Cells(StartRow, 73)
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
                    xSheet.Range("B" & StartRow).NumberFormat = "@"
                    xSheet.Range("B" & StartRow).Value = Right(.Text, 10)
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
