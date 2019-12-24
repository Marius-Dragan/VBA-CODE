Attribute VB_Name = "MMS100MI_LINE"
Option Explicit
Public UpdateUserProfileError As Boolean
Public responseError As String
Sub UploadLineMMS101API()

Call UpdateUserProfile
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
    strProgram = "MMS100MI"
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
    sURL = sURL & "&TRNR=" & lineWS.Cells(StartRow, 3)
    sURL = sURL & "&FACI=" & lineWS.Cells(StartRow, 4)
    sURL = sURL & "&ITNO=" & lineWS.Cells(StartRow, 5)

    If lineWS.Cells(StartRow, 6).Value <> "" Then
    sURL = sURL & "&RSCD=" & lineWS.Cells(StartRow, 6)
    End If

    If lineWS.Cells(StartRow, 7).Value <> "" Then
    sURL = sURL & "&TRQT=" & lineWS.Cells(StartRow, 7)
    End If

    If lineWS.Cells(StartRow, 8).Value <> "" Then
    sURL = sURL & "&WHLO=" & lineWS.Cells(StartRow, 8)
    End If

    If lineWS.Cells(StartRow, 9).Value <> "" Then
    sURL = sURL & "&WHSL=" & lineWS.Cells(StartRow, 9)
    End If

    If lineWS.Cells(StartRow, 10).Value <> "" Then
    sURL = sURL & "&TWSL=" & lineWS.Cells(StartRow, 10)
    End If

    With HttpClient
        .Open "GET", sURL, False, strUsername, strPassword
        .setRequestHeader "Content-Type", "application/xml"
        .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store chache
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

Sub UploadLineMMS101Soap()
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
    strUsername = "INFORBC\" & Sheet1.Cells(1, 2)
    Dim strPassword
    strPassword = Sheet1.Cells(2, 2)
    Dim strServiceRoot
    strServiceRoot = Sheet1.Cells(3, 2)
    Dim strWebServicesServer
    strWebServicesServer = Sheet1.Cells(4, 2)
    Dim strNamespaceBase
    strNamespaceBase = Sheet1.Cells(5, 2)
    Dim strServiceName
    strServiceName = Sheet1.Cells(6, 2)
    Dim strMethod
    strMethod = Sheet1.Cells(7, 2)
    Dim strMethod2
    strMethod2 = Sheet1.Cells(9, 2)

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
