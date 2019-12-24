Attribute VB_Name = "MMS200MI_UPD_ITM_FAC"
Option Explicit
Sub UpdateMMS200MI()

Dim M3Response, M3ResponseName, MSG As String
Dim XmlDoc As Object
Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Dim HttpClient As Object
Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")


'Configure sheet
Dim strTransaction As String
Dim xSheet As Worksheet
Set xSheet = Sheet1
strTransaction = xSheet.Range("B5")

'Configure user details and tranzaction
Dim strUsername
strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
Dim strPassword
strPassword = xSheet.Range("B3")


'Capture the starting Row
Dim StartRow As Long
StartRow = xSheet.Range("B7")
Dim EndRow As Long
EndRow = xSheet.Range("B8")

'Configure environment
Dim strProgram As String
strProgram = "MMS200MI"
Dim sURL As String
Dim envURL As String
If xSheet.Range("B4").Value = "Production" Then
    envURL = "https://yourdomain.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
Else
    envURL = "https://yourdomaindev.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
End If

Application.ScreenUpdating = True
Do While (StartRow <= EndRow)

    sURL = envURL
    sURL = sURL & "&CONO=" & xSheet.Cells(StartRow, 3)
    sURL = sURL & "&FACI=" & xSheet.Cells(StartRow, 4)
    sURL = sURL & "&ITNO=" & xSheet.Cells(StartRow, 5)

    If xSheet.Cells(StartRow, 6).Value <> "" Then
    sURL = sURL & "&LEA4=" & xSheet.Cells(StartRow, 6)
    End If

    If xSheet.Cells(StartRow, 7).Value <> "" Then
    sURL = sURL & "&CSNO=" & xSheet.Cells(StartRow, 7)
    End If

    If xSheet.Cells(StartRow, 8).Value <> "" Then
    sURL = sURL & "&SPFA=" & xSheet.Cells(StartRow, 8)
    End If

    If xSheet.Cells(StartRow, 9).Value <> "" Then
    sURL = sURL & "&ORCO=" & xSheet.Cells(StartRow, 9)
    End If

    If xSheet.Cells(StartRow, 10).Value <> "" Then
    sURL = sURL & "&APPR=" & xSheet.Cells(StartRow, 10)
    End If

    If xSheet.Cells(StartRow, 11).Value <> "" Then
    sURL = sURL & "&UCOS=" & xSheet.Cells(StartRow, 11)
    End If

    If xSheet.Cells(StartRow, 12).Value <> "" Then
    sURL = sURL & "&SOCO=" & xSheet.Cells(StartRow, 12)
    End If

    If xSheet.Cells(StartRow, 13).Value <> "" Then
    sURL = sURL & "&EXPC=" & xSheet.Cells(StartRow, 13)
    End If

    If xSheet.Cells(StartRow, 14).Value <> "" Then
    sURL = sURL & "&BQTY=" & xSheet.Cells(StartRow, 14)
    End If

    If xSheet.Cells(StartRow, 15).Value <> "" Then
    sURL = sURL & "&BQTM=" & xSheet.Cells(StartRow, 15)
    End If

    If xSheet.Cells(StartRow, 16).Value <> "" Then
    sURL = sURL & "&LLCM=" & xSheet.Cells(StartRow, 16)
    End If

    If xSheet.Cells(StartRow, 17).Value <> "" Then
    sURL = sURL & "&DLET=" & xSheet.Cells(StartRow, 17)
    End If

    If xSheet.Cells(StartRow, 18).Value <> "" Then
    sURL = sURL & "&DLEF=" & xSheet.Cells(StartRow, 18)
    End If

    If xSheet.Cells(StartRow, 19).Value <> "" Then
    sURL = sURL & "&DIDY=" & xSheet.Cells(StartRow, 19)
    End If

    If xSheet.Cells(StartRow, 20).Value <> "" Then
    sURL = sURL & "&DIDF=" & xSheet.Cells(StartRow, 20)
    End If

    If xSheet.Cells(StartRow, 21).Value <> "" Then
    sURL = sURL & "&PRRA=" & xSheet.Cells(StartRow, 21)
    End If

    If xSheet.Cells(StartRow, 22).Value <> "" Then
    sURL = sURL & "&TRHC=" & xSheet.Cells(StartRow, 22)
    End If

    If xSheet.Cells(StartRow, 23).Value <> "" Then
    sURL = sURL & "&MARC=" & xSheet.Cells(StartRow, 23)
    End If

    If xSheet.Cells(StartRow, 24).Value <> "" Then
    sURL = sURL & "&JITF=" & xSheet.Cells(StartRow, 24)
    End If

    If xSheet.Cells(StartRow, 25).Value <> "" Then
    sURL = sURL & "&REWH=" & xSheet.Cells(StartRow, 25)
    End If

    If xSheet.Cells(StartRow, 26).Value <> "" Then
    sURL = sURL & "&OPFQ=" & xSheet.Cells(StartRow, 26)
    End If

    If xSheet.Cells(StartRow, 27).Value <> "" Then
    sURL = sURL & "&FANO=" & xSheet.Cells(StartRow, 27)
    End If

    If xSheet.Cells(StartRow, 28).Value <> "" Then
    sURL = sURL & "&FANQ=" & xSheet.Cells(StartRow, 28)
    End If

    If xSheet.Cells(StartRow, 29).Value <> "" Then
    sURL = sURL & "&FANR=" & xSheet.Cells(StartRow, 29)
    End If

    If xSheet.Cells(StartRow, 30).Value <> "" Then
    sURL = sURL & "&FATM=" & xSheet.Cells(StartRow, 30)
    End If

    If xSheet.Cells(StartRow, 31).Value <> "" Then
    sURL = sURL & "&WCLN=" & xSheet.Cells(StartRow, 31)
    End If

    If xSheet.Cells(StartRow, 32).Value <> "" Then
    sURL = sURL & "&EDEC=" & xSheet.Cells(StartRow, 32)
    End If

    If xSheet.Cells(StartRow, 33).Value <> "" Then
    sURL = sURL & "&AUGE=" & xSheet.Cells(StartRow, 33)
    End If

    If xSheet.Cells(StartRow, 34).Value <> "" Then
    sURL = sURL & "&ECCC=" & xSheet.Cells(StartRow, 34)
    End If

    If xSheet.Cells(StartRow, 35).Value <> "" Then
    sURL = sURL & "&ECAR=" & xSheet.Cells(StartRow, 35)
    End If

    If xSheet.Cells(StartRow, 36).Value <> "" Then
    sURL = sURL & "&CPRI=" & xSheet.Cells(StartRow, 36)
    End If

    If xSheet.Cells(StartRow, 37).Value <> "" Then
    sURL = sURL & "&CPRE=" & xSheet.Cells(StartRow, 37)
    End If

    If xSheet.Cells(StartRow, 38).Value <> "" Then
    sURL = sURL & "&WSCA=" & xSheet.Cells(StartRow, 38)
    End If

    If xSheet.Cells(StartRow, 39).Value <> "" Then
    sURL = sURL & "&PRCM=" & xSheet.Cells(StartRow, 39)
    End If

    If xSheet.Cells(StartRow, 40).Value <> "" Then
    sURL = sURL & "&PLAP=" & xSheet.Cells(StartRow, 40)
    End If

    If xSheet.Cells(StartRow, 41).Value <> "" Then
    sURL = sURL & "&PLUP=" & xSheet.Cells(StartRow, 41)
    End If

    If xSheet.Cells(StartRow, 42).Value <> "" Then
    sURL = sURL & "&SCMO=" & xSheet.Cells(StartRow, 42)
    End If

    If xSheet.Cells(StartRow, 43).Value <> "" Then
    sURL = sURL & "&CPL0=" & xSheet.Cells(StartRow, 43)
    End If

    If xSheet.Cells(StartRow, 44).Value <> "" Then
    sURL = sURL & "&CPL1=" & xSheet.Cells(StartRow, 44)
    End If

    If xSheet.Cells(StartRow, 45).Value <> "" Then
    sURL = sURL & "&CPL2=" & xSheet.Cells(StartRow, 45)
    End If

    If xSheet.Cells(StartRow, 46).Value <> "" Then
    sURL = sURL & "&PPL0=" & xSheet.Cells(StartRow, 46)
    End If

    If xSheet.Cells(StartRow, 47).Value <> "" Then
    sURL = sURL & "&PPL1=" & xSheet.Cells(StartRow, 47)
    End If

    If xSheet.Cells(StartRow, 48).Value <> "" Then
    sURL = sURL & "&PPL2=" & xSheet.Cells(StartRow, 48)
    End If

    If xSheet.Cells(StartRow, 49).Value <> "" Then
    sURL = sURL & "&TXID=" & xSheet.Cells(StartRow, 49)
    End If

    If xSheet.Cells(StartRow, 50).Value <> "" Then
    sURL = sURL & "&DTID=" & xSheet.Cells(StartRow, 50)
    End If

    If xSheet.Cells(StartRow, 51).Value <> "" Then
    sURL = sURL & "&CPDC=" & xSheet.Cells(StartRow, 51)
    End If

    If xSheet.Cells(StartRow, 52).Value <> "" Then
    sURL = sURL & "&COCD=" & xSheet.Cells(StartRow, 52)
    End If

    If xSheet.Cells(StartRow, 53).Value <> "" Then
    sURL = sURL & "&EVGR=" & xSheet.Cells(StartRow, 53)
    End If

    If xSheet.Cells(StartRow, 54).Value <> "" Then
    sURL = sURL & "&VAMT=" & xSheet.Cells(StartRow, 54)
    End If

    If xSheet.Cells(StartRow, 55).Value <> "" Then
    sURL = sURL & "&LAMA=" & xSheet.Cells(StartRow, 55)
    End If

    If xSheet.Cells(StartRow, 56).Value <> "" Then
    sURL = sURL & "&GRTI=" & xSheet.Cells(StartRow, 56)
    End If

    If xSheet.Cells(StartRow, 57).Value <> "" Then
    sURL = sURL & "&MOLL=" & xSheet.Cells(StartRow, 57)
    End If

    If xSheet.Cells(StartRow, 58).Value <> "" Then
    sURL = sURL & "&CRTM=" & xSheet.Cells(StartRow, 58)
    End If

    If xSheet.Cells(StartRow, 59).Value <> "" Then
    sURL = sURL & "&DICM=" & xSheet.Cells(StartRow, 59)
    End If

    If xSheet.Cells(StartRow, 60).Value <> "" Then
    sURL = sURL & "&ACRF=" & xSheet.Cells(StartRow, 60)
    End If

    If xSheet.Cells(StartRow, 61).Value <> "" Then
    sURL = sURL & "&STCW=" & xSheet.Cells(StartRow, 61)
    End If

    If xSheet.Cells(StartRow, 62).Value <> "" Then
    sURL = sURL & "&RJCW=" & xSheet.Cells(StartRow, 62)
    End If

    If xSheet.Cells(StartRow, 63).Value <> "" Then
    sURL = sURL & "&QUCW=" & xSheet.Cells(StartRow, 63)
    End If

    If xSheet.Cells(StartRow, 64).Value <> "" Then
    sURL = sURL & "&CAWC=" & xSheet.Cells(StartRow, 64)
    End If

    If xSheet.Cells(StartRow, 65).Value <> "" Then
    sURL = sURL & "&CPUN=" & xSheet.Cells(StartRow, 65)
    End If

    If xSheet.Cells(StartRow, 66).Value <> "" Then
    sURL = sURL & "&COFA=" & xSheet.Cells(StartRow, 66)
    End If

    If xSheet.Cells(StartRow, 67).Value <> "" Then
    sURL = sURL & "&ALTS=" & xSheet.Cells(StartRow, 67)
    End If

    If xSheet.Cells(StartRow, 68).Value <> "" Then
    sURL = sURL & "&ATTC=" & xSheet.Cells(StartRow, 68)
    End If


    With HttpClient
        .Open "GET", sURL, False, strUsername, strPassword
        .setRequestHeader "Content-Type", "application/xml"
        .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store cache
        .setRequestHeader "Authorization", "Basic " + Encoding.Base64Encode(strUsername + ":" + strPassword)
        .send 'send HTTP request
        M3Response = .responseText
    End With

    'Debug.Print M3Response
    With XmlDoc
        .LoadXML M3Response
        M3ResponseName = .DocumentElement.nodeName
        If M3ResponseName = "ErrorMessage" Then

            MSG = .DocumentElement.FirstChild.Text
          'Debug.Print msg
          xSheet.Range("A" & StartRow).Value = "NOK"
          xSheet.Range("B" & StartRow).Value = MSG
          xSheet.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
          xSheet.Range("B" & StartRow).Replace "  ", ""
        Else
            MSG = .DocumentElement.FirstChild.Text
            xSheet.Range("A" & StartRow).Value = "OK"
            'xSheet.Range("B" & StartRow).Value = MSG & " Uploaded OK"
        End If
    End With

    sURL = ""
    StartRow = StartRow + 1

Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, strProgram & " " & strTransaction
End Sub
Sub ClearHeaderlogs()
Dim LRow As Long
    With Sheet1
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
