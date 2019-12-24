Attribute VB_Name = "PPS200MI_LINE"
Option Explicit
Public UpdateUserProfileError As Boolean
Public responseError As String
Sub Upload_PO_PPS200_Line()

Call UserM3ProfileChange.UpdateUserProfile
If UpdateUserProfileError = False Then

    Dim M3Response, M3ResponseName, MSG As String
    Dim XmlDoc As Object
    Dim HttpClient As Object
    Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Set HttpClient = CreateObject("MSXML2.XMLHTTP.6.0")

    'Configure sheet and tranzaction
    Dim strTransaction As String
    Dim lineWS As Worksheet
    Set lineWS = Sheet2
    strTransaction = lineWS.Range("B5")

    'Configure settings
    Dim strUsername
    strUsername = "INFORBC\" & UCase(lineWS.Range("B2"))
    Dim strPassword
    strPassword = lineWS.Range("B3")

    'Configure env
    Dim strProgram As String
    strProgram = "PPS200MI"
    Dim sURL As String
    Dim envURL As String
    If lineWS.Range("B4").Value = "Production" Then
        envURL = "https://yourdomain.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
    Else
        envURL = "https://yourdomaindev.cloud.infor.com:12345/m3api-rest/execute/" & strProgram & "/" & strTransaction & "?"
    End If

    'Capture the starting Row
    Dim StartRow As Long
    StartRow = lineWS.Range("B7")
    Dim EndRow As Long
    EndRow = lineWS.Range("B8")

    Application.ScreenUpdating = False
    Do While (StartRow <= EndRow)

    sURL = envURL
    sURL = sURL & "&PUNO=" & lineWS.Cells(StartRow, 3)

    If lineWS.Cells(StartRow, 4).Value <> vbNullString Then
        sURL = sURL & "&PNLI=" & lineWS.Cells(StartRow, 4)
    End If

    If lineWS.Cells(StartRow, 5).Value <> vbNullString Then
        sURL = sURL & "&PNLS=" & lineWS.Cells(StartRow, 5)
    End If

    'If lineWS.Cells(StartRow, 6).Value <> vbNullString Then
        sURL = sURL & "&ITNO=" & lineWS.Cells(StartRow, 6)
    'End If

    'If lineWS.Cells(StartRow, 7).Value <> vbNullString Then
        sURL = sURL & "&ORQA=" & lineWS.Cells(StartRow, 7)
    'End If

    If lineWS.Cells(StartRow, 8).Value <> vbNullString Then
        sURL = sURL & "&FACI=" & lineWS.Cells(StartRow, 8)
    End If

    'If lineWS.Cells(StartRow, 9).Value <> vbNullString Then
        sURL = sURL & "&WHLO=" & lineWS.Cells(StartRow, 9)
    'End If

    If lineWS.Cells(StartRow, 10).Value <> vbNullString Then
        sURL = sURL & "&SUNO=" & lineWS.Cells(StartRow, 10)
    End If

    If lineWS.Cells(StartRow, 11).Value <> vbNullString Then
        sURL = sURL & "&DWDT=" & lineWS.Cells(StartRow, 11)
    End If

    'If lineWS.Cells(StartRow, 12).Value <> vbNullString Then
        sURL = sURL & "&SITE=" & lineWS.Cells(StartRow, 12)
    'End If

    If lineWS.Cells(StartRow, 13).Value <> vbNullString Then
        sURL = sURL & "&PITD=" & lineWS.Cells(StartRow, 13)
    End If

    If lineWS.Cells(StartRow, 14).Value <> vbNullString Then
        sURL = sURL & "&PITT=" & lineWS.Cells(StartRow, 14)
    End If

    If lineWS.Cells(StartRow, 15).Value <> vbNullString Then
        sURL = sURL & "&PROD=" & lineWS.Cells(StartRow, 15)
    End If

    If lineWS.Cells(StartRow, 16).Value <> vbNullString Then
        sURL = sURL & "&ECVE=" & lineWS.Cells(StartRow, 16)
    End If

    If lineWS.Cells(StartRow, 17).Value <> vbNullString Then
        sURL = sURL & "&REVN=" & lineWS.Cells(StartRow, 17)
    End If

    If lineWS.Cells(StartRow, 18).Value <> vbNullString Then
        sURL = sURL & "&ETRF=" & lineWS.Cells(StartRow, 18)
    End If

    If lineWS.Cells(StartRow, 19).Value <> vbNullString Then
        sURL = sURL & "&PUPR=" & lineWS.Cells(StartRow, 19)
    End If

    If lineWS.Cells(StartRow, 20).Value <> vbNullString Then
        sURL = sURL & "&ODI1=" & lineWS.Cells(StartRow, 20)
    End If

    If lineWS.Cells(StartRow, 21).Value <> vbNullString Then
        sURL = sURL & "&ODI2=" & lineWS.Cells(StartRow, 21)
    End If

    If lineWS.Cells(StartRow, 22).Value <> vbNullString Then
        sURL = sURL & "&ODI3=" & lineWS.Cells(StartRow, 22)
    End If

    If lineWS.Cells(StartRow, 23).Value <> vbNullString Then
        sURL = sURL & "&PUUN=" & lineWS.Cells(StartRow, 23)
    End If

    If lineWS.Cells(StartRow, 24).Value <> vbNullString Then
        sURL = sURL & "&PPUN=" & lineWS.Cells(StartRow, 24)
    End If

    If lineWS.Cells(StartRow, 25).Value <> vbNullString Then
        sURL = sURL & "&PUCD=" & lineWS.Cells(StartRow, 25)
    End If

    If lineWS.Cells(StartRow, 26).Value <> vbNullString Then
        sURL = sURL & "&PTCD=" & lineWS.Cells(StartRow, 26)
    End If

    If lineWS.Cells(StartRow, 27).Value <> vbNullString Then
        sURL = sURL & "&RORC=" & lineWS.Cells(StartRow, 27)
    End If

    If lineWS.Cells(StartRow, 28).Value <> vbNullString Then
        sURL = sURL & "&RORN=" & lineWS.Cells(StartRow, 28)
    End If

    If lineWS.Cells(StartRow, 29).Value <> vbNullString Then
        sURL = sURL & "&RORL=" & lineWS.Cells(StartRow, 29)
    End If

    If lineWS.Cells(StartRow, 30).Value <> vbNullString Then
        sURL = sURL & "&RORX=" & lineWS.Cells(StartRow, 30)
    End If

    If lineWS.Cells(StartRow, 31).Value <> vbNullString Then
        sURL = sURL & "&OURR=" & lineWS.Cells(StartRow, 31)
    End If

    If lineWS.Cells(StartRow, 32).Value <> vbNullString Then
        sURL = sURL & "&OURT=" & lineWS.Cells(StartRow, 32)
    End If

    If lineWS.Cells(StartRow, 33).Value <> vbNullString Then
        sURL = sURL & "&PRIP=" & lineWS.Cells(StartRow, 33)
    End If

    If lineWS.Cells(StartRow, 34).Value <> vbNullString Then
        sURL = sURL & "&FUSC=" & lineWS.Cells(StartRow, 34)
    End If

    If lineWS.Cells(StartRow, 35).Value <> vbNullString Then
        sURL = sURL & "&PURC=" & lineWS.Cells(StartRow, 35)
    End If

    If lineWS.Cells(StartRow, 36).Value <> vbNullString Then
        sURL = sURL & "&BUYE=" & lineWS.Cells(StartRow, 36)
    End If

    If lineWS.Cells(StartRow, 37).Value <> vbNullString Then
        sURL = sURL & "&TERE=" & lineWS.Cells(StartRow, 37)
    End If

    If lineWS.Cells(StartRow, 38).Value <> vbNullString Then
        sURL = sURL & "&GRMT=" & lineWS.Cells(StartRow, 38)
    End If

    If lineWS.Cells(StartRow, 39).Value <> vbNullString Then
        sURL = sURL & "&IRCV=" & lineWS.Cells(StartRow, 39)
    End If

    If lineWS.Cells(StartRow, 40).Value <> vbNullString Then
        sURL = sURL & "&PACT=" & lineWS.Cells(StartRow, 40)
    End If

    If lineWS.Cells(StartRow, 41).Value <> vbNullString Then
        sURL = sURL & "&VTCD=" & lineWS.Cells(StartRow, 41)
    End If

    If lineWS.Cells(StartRow, 42).Value <> vbNullString Then
        sURL = sURL & "&ACRF=" & lineWS.Cells(StartRow, 42)
    End If

    If lineWS.Cells(StartRow, 43).Value <> vbNullString Then
        sURL = sURL & "&COCE=" & lineWS.Cells(StartRow, 43)
    End If

    If lineWS.Cells(StartRow, 44).Value <> vbNullString Then
        sURL = sURL & "&CSNO=" & lineWS.Cells(StartRow, 44)
    End If

    If lineWS.Cells(StartRow, 45).Value <> vbNullString Then
        sURL = sURL & "&ECLC=" & lineWS.Cells(StartRow, 45)
    End If

    If lineWS.Cells(StartRow, 46).Value <> vbNullString Then
        sURL = sURL & "&VRCD=" & lineWS.Cells(StartRow, 46)
    End If

    If lineWS.Cells(StartRow, 47).Value <> vbNullString Then
        sURL = sURL & "&PROJ=" & lineWS.Cells(StartRow, 47)
    End If

    If lineWS.Cells(StartRow, 48).Value <> vbNullString Then
        sURL = sURL & "&ELNO=" & lineWS.Cells(StartRow, 48)
    End If

    If lineWS.Cells(StartRow, 49).Value <> vbNullString Then
        sURL = sURL & "&CPRI=" & lineWS.Cells(StartRow, 49)
    End If

    If lineWS.Cells(StartRow, 50).Value <> vbNullString Then
        sURL = sURL & "&HAFE=" & lineWS.Cells(StartRow, 50)
    End If

    If lineWS.Cells(StartRow, 51).Value <> vbNullString Then
        sURL = sURL & "&TAXC=" & lineWS.Cells(StartRow, 51)
    End If

    If lineWS.Cells(StartRow, 52).Value <> vbNullString Then
        sURL = sURL & "&TIHM=" & lineWS.Cells(StartRow, 52)
    End If

    If lineWS.Cells(StartRow, 53).Value <> vbNullString Then
        sURL = sURL & "&MSTN=" & lineWS.Cells(StartRow, 53)
    End If

    If lineWS.Cells(StartRow, 54).Value <> vbNullString Then
        sURL = sURL & "&UPCK=" & lineWS.Cells(StartRow, 54)
    End If

    If lineWS.Cells(StartRow, 55).Value <> vbNullString Then
        sURL = sURL & "&ORCO=" & lineWS.Cells(StartRow, 55)
    End If

    If lineWS.Cells(StartRow, 56).Value <> vbNullString Then
        sURL = sURL & "&GEOC=" & lineWS.Cells(StartRow, 56)
    End If

    If lineWS.Cells(StartRow, 57).Value <> vbNullString Then
        sURL = sURL & "&TRRC=" & lineWS.Cells(StartRow, 57)
    End If

    If lineWS.Cells(StartRow, 58).Value <> vbNullString Then
        sURL = sURL & "&TRRN=" & lineWS.Cells(StartRow, 58)
    End If

    If lineWS.Cells(StartRow, 59).Value <> vbNullString Then
        sURL = sURL & "&TRRL=" & lineWS.Cells(StartRow, 59)
    End If

    If lineWS.Cells(StartRow, 60).Value <> vbNullString Then
        sURL = sURL & "&TRRX=" & lineWS.Cells(StartRow, 60)
    End If

    If lineWS.Cells(StartRow, 61).Value <> vbNullString Then
        sURL = sURL & "&RASN=" & lineWS.Cells(StartRow, 61)
    End If

    If lineWS.Cells(StartRow, 62).Value <> vbNullString Then
        sURL = sURL & "&PIAD=" & lineWS.Cells(StartRow, 62)
    End If

    If lineWS.Cells(StartRow, 63).Value <> vbNullString Then
        sURL = sURL & "&ORAD=" & lineWS.Cells(StartRow, 63)
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
                    lineWS.Range("A" & StartRow).Value = "OK"
                    'lineWS.Range("B" & StartRow).Value = MSG & " Updated OK"
                    lineWS.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                    lineWS.Range("B" & StartRow).Replace "  ", ""
                    End With
                Else ' Failure (404, 500, ...)
                With XmlDoc
                        MSG = .DocumentElement.FirstChild.Text
                      'Debug.Print msg
                      lineWS.Range("A" & StartRow).Value = "NOK"
                      lineWS.Range("B" & StartRow).Value = MSG
                      lineWS.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                      lineWS.Range("B" & StartRow).Replace "  ", ""
                  End With
            End If
    Else
        MsgBox "Error: " & HttpClient.Status & " " & HttpClient.statusText, vbCritical, "MNS150MI" & strTransaction
    Exit Sub
    End If

    sURL = ""
    StartRow = StartRow + 1

    Loop
        Application.ScreenUpdating = True
        MsgBox "Process completed!", vbInformation, strProgram & " " & strTransaction
Else
        MsgBox "Please check M3 user profile settings and try again." & vbNewLine & "Error: " & responseError, vbInformation, "M3 Profile Settings Update Error"
End If
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

Private Sub UploadLineMMS101Soap()
'Not in use needs updating
Call UpdateUserProfile
If UpdateUserProfileError = False Then

    Dim responseText As String
    Dim sURL As String
    Dim sEnv As String
    Dim xmlhtp As New MSXML2.XMLHTTP60
    Dim XmlDoc As New DOMDocument60
    Dim XmlNode As MSXML2.IXMLDOMNode
    Dim XmlNodeList As MSXML2.IXMLDOMNodeList
    Dim MyNode As MSXML2.IXMLDOMNode

    ' Settings Definition
    Dim strUsername
    strUsername = "INFORBC\" & lineWS.Cells(1, 2)
    Dim strPassword
    strPassword = lineWS.Cells(2, 2)
    Dim strServiceRoot
    strServiceRoot = lineWS.Cells(3, 2)
    Dim strWebServicesServer
    strWebServicesServer = lineWS.Cells(4, 2)
    Dim strNamespaceBase
    strNamespaceBase = lineWS.Cells(5, 2)
    Dim strServiceName
    strServiceName = lineWS.Cells(6, 2)
    Dim strMethod
    strMethod = lineWS.Cells(7, 2)
    Dim strMethod2
    strMethod2 = lineWS.Cells(9, 2)

    Dim strTargetNamespace
    strTargetNamespace = strNamespaceBase & "/" & strServiceName & "/" & strMethod2

    Dim strSoapAction
    strSoapAction = strWebServicesServer & "/" & strServiceName

    sURL = strServiceRoot & "/" & strServiceName

    Dim StartRow As Long
    StartRow = Sheet3.Cells(5, 1)
    Dim EndRow As Long
    EndRow = Sheet3.Cells(7, 1)

    Application.ScreenUpdating = False
    Do While (StartRow <= EndRow)
        ' Clear Return Code and Error Message
        Sheet3.Cells(StartRow, 2).ClearContents
        Sheet3.Cells(StartRow, 3).ClearContents

        'Add Controls to terminate the Loop when no information in the Cell!
        If IsEmpty(Sheet3.Cells(StartRow, 6)) = True Then
            'Goodbye!
            Exit Do
        Else

        sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:cred=""http://lawson.com/ws/credentials"" xmlns:mms=" & """" & strTargetNamespace & """" & ">"
        sEnv = sEnv & "<soapenv:Header>"
        sEnv = sEnv & " <cred:lws>"
        sEnv = sEnv & "   <!--Optional:-->"
        sEnv = sEnv & "    <cred:company>?</cred:company>"
        sEnv = sEnv & "    <!--Optional:-->"
        sEnv = sEnv & "     <cred:division>?</cred:division>"
        sEnv = sEnv & "  </cred:lws>"
        sEnv = sEnv & " </soapenv:Header>"
        sEnv = sEnv & " <soapenv:Body>"
        sEnv = sEnv & "   <mms:MMS100_Lines>"
        sEnv = sEnv & "     <mms:MMS100>"
            If IsEmpty(Sheet3.Cells(StartRow, 5)) = False Then
            sEnv = sEnv & "       <mms:OrderNumber>" & Sheet3.Cells(StartRow, 5) & "</mms:OrderNumber>"
            End If
            sEnv = sEnv & "      <mms:Facility>" & Sheet3.Cells(StartRow, 6) & "</mms:Facility>"
            sEnv = sEnv & "      <mms:MMS101>"
                sEnv = sEnv & "          <mms:ItemNumber>" & Sheet3.Cells(StartRow, 7) & "</mms:ItemNumber>"
                sEnv = sEnv & "          <mms:TransactionQuantityBasicUM>" & Sheet3.Cells(StartRow, 9) & "</mms:TransactionQuantityBasicUM>"
                sEnv = sEnv & "          <mms:Warehouse>" & Sheet3.Cells(StartRow, 10) & "</mms:Warehouse>"
                If IsEmpty(Sheet3.Cells(StartRow, 11)) = False Then
                    sEnv = sEnv & "          <mms:Location>" & Sheet3.Cells(StartRow, 11) & "</mms:Location>"
                End If
                If IsEmpty(Sheet3.Cells(StartRow, 8)) = False Then
                    sEnv = sEnv & "          <mms:TransactionReason>" & Sheet3.Cells(StartRow, 8) & "</mms:TransactionReason>"
                End If
            sEnv = sEnv & "       <mms:MMS101>"
        sEnv = sEnv & "     </mms:MMS100>"
        sEnv = sEnv & "   </mms:MMS100_Lines>"
        sEnv = sEnv & "</soapenv:Body>"
        sEnv = sEnv & "</soapenv:Envelope>"

MsgBox (sEnv)

            With xmlhtp

                .Open "POST", sURL, False, strUsername, strPassword
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", strSoapAction
                .setRequestHeader "Authorization", "Basic " + Base64Encode(strUsername + ":" + strPassword)
                .send sEnv

                XmlDoc.LoadXML (.responseText)

                If xmlhtp.Status = 200 Then
                    ' Success (200)
                    'MsgBox "OK " & xmlhtp.Status
                    Sheet3.Cells(StartRow, 2).Value = "OK"
                Else
                    ' Failure (404, 500, ...)
                    'MsgBox "NOK " & xmlhtp.Status
                    Sheet3.Cells(StartRow, 2).Value = "NOK"
                    Set XmlNodeList = XmlDoc.getElementsByTagName("*")
                    For Each XmlNode In XmlNodeList
                        For Each MyNode In XmlNode.ChildNodes
                            If MyNode.NodeType = NODE_TEXT Then
                                If XmlNode.nodeName = "faultstring" Then
                                    'MsgBox xmlNode.nodeName & "=" & xmlNode.Text
                                    Sheet3.Cells(StartRow, 3).Value = XmlNode.Text
                                End If
                            End If
                        Next MyNode
                    Next XmlNode
                End If
                'MsgBox .responseText
            End With
        End If
        ' Continue with next Row
        StartRow = StartRow + 1
    Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, "M3 Upload"

Else
    MsgBox "Please check M3 user profile settings and try again." & vbNewLine & "Error: " & responseError, vbInformation, "M3 Profile Settings Update Error"
End If
End Sub
