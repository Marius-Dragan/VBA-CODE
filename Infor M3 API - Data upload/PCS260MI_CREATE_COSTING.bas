Attribute VB_Name = "PCS260MI_CREATE_COSTING"
Option Explicit
Public UpdateUserProfileError As Boolean
Public responseError As String
Sub UploadHeaderPCS260()

Call UserM3ProfileChange.UpdateUserProfile
If UpdateUserProfileError = False Then

    'Early Binding requires reference added "Microsoft XML v6"
    '    Dim xmlhtp As New MSXML2.XMLHTTP60
    '    Dim XmlDoc As New DOMDocument60
    '    Dim xmlNode As MSXML2.IXMLDOMNode
    '    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    '    Dim myNode As MSXML2.IXMLDOMNode

    'Late Binding no dependencies
    Dim XmlHttp As Object
    Set XmlHttp = CreateObject("MSXML2.XMLHTTP")
    Dim XmlDoc As Object
    Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Dim XmlNode As Object
    Dim XmlNodeList As Object

    'Configure sheet and tranzaction
    Dim xSheet As Worksheet
    Dim strTransaction As String
    Set xSheet = Sheet1
    strTransaction = xSheet.Range("B5")

    ' Config details
    Dim strUsername
    strUsername = "INFORBC\" & UCase(xSheet.Range("B2"))
    Dim strPassword
    strPassword = xSheet.Range("B3")
    Dim StartRow As Long
    StartRow = xSheet.Range("B7")
    Dim EndRow As Long
    EndRow = xSheet.Range("B8")

    'Configure services
    Dim sURL As String
    Dim sEnv As String
    Dim strServiceRoot
    Dim strWebServicesServer
    If xSheet.Range("B4").Value = "Production" Then
        strServiceRoot = "https://yourdomain.cloud.infor.com:12345/mws-ws/services"
        strWebServicesServer = "https://yourdomain.cloud.infor.com"
    Else
        strServiceRoot = "https://yourdomaindev.cloud.infor.com:12345/mws-ws/services"
        strWebServicesServer = "https://yourdomaindev.cloud.infor.com"
    End If

    Dim strNamespaceBase
    strNamespaceBase = "http://your.company.net"
    Dim strServiceName
    strServiceName = "PCS260_CREATE"
    Dim strMethod
    strMethod = "PCS260_CREATECOSTING"

    Dim strTargetNamespace
    strTargetNamespace = strNamespaceBase & "/" & strServiceName & "/" & strMethod

    Dim strSoapAction
    strSoapAction = strWebServicesServer & "/" & strServiceName

    sURL = strServiceRoot & "/" & strServiceName

    Application.ScreenUpdating = False

    Do While (StartRow <= EndRow)

        sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:cred=""http://lawson.com/ws/credentials"" xmlns:pcs=" & """" & strTargetNamespace & """" & ">"
        sEnv = sEnv & "   <soapenv:Header>"
        sEnv = sEnv & "      <cred:lws>"
        sEnv = sEnv & "         <!--Optional:-->"
        sEnv = sEnv & "         <cred:company>?</cred:company>"
        sEnv = sEnv & "         <!--Optional:-->"
        sEnv = sEnv & "         <cred:division>?</cred:division>"
        sEnv = sEnv & "      </cred:lws>"
        sEnv = sEnv & "   </soapenv:Header>"
        sEnv = sEnv & "   <soapenv:Body>"
        sEnv = sEnv & "      <pcs:PCS260_CREATECOSTING>"
        sEnv = sEnv & "         <pcs:PCS260>"
        sEnv = sEnv & "            <pcs:Facility>" & xSheet.Cells(StartRow, 3) & "</pcs:Facility>"
        sEnv = sEnv & "            <pcs:ItemNumber>" & xSheet.Cells(StartRow, 4) & "</pcs:ItemNumber>"
        sEnv = sEnv & "            <pcs:ProductStructureType>" & xSheet.Cells(StartRow, 5) & "</pcs:ProductStructureType>"

        If IsEmpty(xSheet.Cells(StartRow, 6)) = False Then
            sEnv = sEnv & "            <pcs:ConfigID>" & xSheet.Cells(StartRow, 6) & "</pcs:ConfigID>"
        End If

        If IsEmpty(xSheet.Cells(StartRow, 7)) = False Then
            sEnv = sEnv & "            <pcs:RefOrderNo>" & xSheet.Cells(StartRow, 7) & "</pcs:RefOrderNo>"
        End If

        If IsEmpty(xSheet.Cells(StartRow, 8)) = False Then
            sEnv = sEnv & "            <pcs:CostingType>" & xSheet.Cells(StartRow, 8) & "</pcs:CostingType>"
        End If

        sEnv = sEnv & "            <pcs:CostingDate>" & xSheet.Cells(StartRow, 9) & "</pcs:CostingDate>"
        sEnv = sEnv & "            <pcs:CostingSum1>" & xSheet.Cells(StartRow, 10) & "</pcs:CostingSum1>"
        sEnv = sEnv & "         </pcs:PCS260>"
        sEnv = sEnv & "      </pcs:PCS260_CREATECOSTING>"
        sEnv = sEnv & "   </soapenv:Body>"
        sEnv = sEnv & "</soapenv:Envelope>"


         With XmlHttp

             .Open "POST", sURL, False, strUsername, strPassword
             .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
             .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store cache
             .setRequestHeader "soapAction", strSoapAction
             .setRequestHeader "Authorization", "Basic " + Encoding.Base64Encode(strUsername + ":" + strPassword)
             .send sEnv

             XmlDoc.LoadXML (.responseText)

                If XmlHttp.Status = 200 Then ' Success (200)
                    xSheet.Cells(StartRow, 1).Value = "OK"
                Else ' Failure (404, 500, ...)
                    xSheet.Cells(StartRow, 1).Value = "NOK"
                    Set XmlNodeList = XmlDoc.getElementsByTagName("*")
                    For Each XmlNode In XmlNodeList
                        If XmlNode.nodeName = "faultstring" Then
                            xSheet.Cells(StartRow, 2).Value = XmlNode.Text
                        End If
                    Next XmlNode
                End If
        End With
        ' Continue with next Row
        StartRow = StartRow + 1
    Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, strMethod
Else
    Application.ScreenUpdating = True
    MsgBox "Please check M3 user profile settings and try again." & vbNewLine & "Error: " & responseError, vbInformation, "M3 Profile Settings Update Error"
End If

End Sub
Sub ClearLogsHead()
Dim LRow As Long
    With Sheet1
        LRow = .Range("A" & .Rows.Count).End(xlUp).Row
        If LRow = 14 Then
            LRow = LRow + 1
        End If
        .Range("A15:B" & LRow).ClearContents
    End With
End Sub
