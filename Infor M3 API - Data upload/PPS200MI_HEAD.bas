Attribute VB_Name = "PPS200MI_HEAD"
Option Explicit
Public UpdateUserProfileError As Boolean
Public responseError As String
Sub Upload_PO_PPS200_Head()

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
    strServiceName = "PurchaseOrders"
    Dim strMethod
    strMethod = "PPS200"

    Dim strTargetNamespace
    strTargetNamespace = strNamespaceBase & "/" & strServiceName & "/" & strMethod

    Dim strSoapAction
    strSoapAction = strWebServicesServer & "/" & strServiceName

    sURL = strServiceRoot & "/" & strServiceName

    Application.ScreenUpdating = False

    Do While (StartRow <= EndRow)
            sEnv = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:cred=""http://lawson.com/ws/credentials"" xmlns:pps=" & """" & strTargetNamespace & """" & ">"
            sEnv = sEnv & "   <soapenv:Header>"
            sEnv = sEnv & "      <cred:lws>"
            sEnv = sEnv & "         <!--Optional:-->"
            sEnv = sEnv & "         <cred:company>?</cred:company>"
            sEnv = sEnv & "         <!--Optional:-->"
            sEnv = sEnv & "         <cred:division>?</cred:division>"
            sEnv = sEnv & "      </cred:lws>"
            sEnv = sEnv & "   </soapenv:Header>"
            sEnv = sEnv & "   <soapenv:Body>"
            sEnv = sEnv & "      <pps:PPS200>"
            sEnv = sEnv & "         <pps:PPS200>"
            sEnv = sEnv & "            <pps:Facility>" & xSheet.Cells(StartRow, 3) & "</pps:Facility>"
            sEnv = sEnv & "            <pps:Warehouse>" & xSheet.Cells(StartRow, 4) & "</pps:Warehouse>"
            sEnv = sEnv & "            <pps:PurchaseOrderNumber>" & xSheet.Cells(StartRow, 5) & "</pps:PurchaseOrderNumber>"
            sEnv = sEnv & "            <pps:Supplier>" & xSheet.Cells(StartRow, 6) & "</pps:Supplier>"
            sEnv = sEnv & "            <pps:RequestedDeliveryDate>" & xSheet.Cells(StartRow, 7) & "</pps:RequestedDeliveryDate>"
            sEnv = sEnv & "            <pps:OrderType>" & xSheet.Cells(StartRow, 8) & "</pps:OrderType>"
            sEnv = sEnv & "            <pps:AgreementNumber>" & xSheet.Cells(StartRow, 9) & "</pps:AgreementNumber>"
            sEnv = sEnv & "            <pps:Buyer>" & xSheet.Cells(StartRow, 10) & "</pps:Buyer>"
            sEnv = sEnv & "            <pps:RequisitionBy>" & xSheet.Cells(StartRow, 11) & "</pps:RequisitionBy>"
            sEnv = sEnv & "            <pps:OrderDate>" & xSheet.Cells(StartRow, 12) & "</pps:OrderDate>"
            sEnv = sEnv & "            <pps:DeliveryTerms>" & xSheet.Cells(StartRow, 13) & "</pps:DeliveryTerms>"
            sEnv = sEnv & "            <pps:DeliveryMethod>" & xSheet.Cells(StartRow, 14) & "</pps:DeliveryMethod>"
            sEnv = sEnv & "            <pps:FreightTerms>" & xSheet.Cells(StartRow, 15) & "</pps:FreightTerms>"
            sEnv = sEnv & "            <pps:HarborOrAirport>" & xSheet.Cells(StartRow, 16) & "</pps:HarborOrAirport>"
            sEnv = sEnv & "            <pps:PackagingTerms>" & xSheet.Cells(StartRow, 17) & "</pps:PackagingTerms>"
            sEnv = sEnv & "            <pps:RailStation>" & xSheet.Cells(StartRow, 18) & "</pps:RailStation>"
            sEnv = sEnv & "            <pps:PaymentTerms>" & xSheet.Cells(StartRow, 19) & "</pps:PaymentTerms>"
            sEnv = sEnv & "            <pps:MonitoringActivityList>" & xSheet.Cells(StartRow, 20) & "</pps:MonitoringActivityList>"
            sEnv = sEnv & "            <pps:Currency>" & xSheet.Cells(StartRow, 21) & "</pps:Currency>"
            sEnv = sEnv & "            <pps:PaymentMethodAccountsPayable>" & xSheet.Cells(StartRow, 22) & "</pps:PaymentMethodAccountsPayable>"
            sEnv = sEnv & "            <pps:Language>" & xSheet.Cells(StartRow, 23) & "</pps:Language>"
            sEnv = sEnv & "            <pps:OurReferenceNumber>" & xSheet.Cells(StartRow, 24) & "</pps:OurReferenceNumber>"
            sEnv = sEnv & "            <pps:ReferenceType>" & xSheet.Cells(StartRow, 25) & "</pps:ReferenceType>"
            sEnv = sEnv & "            <pps:YourReference>" & xSheet.Cells(StartRow, 26) & "</pps:YourReference>"
            sEnv = sEnv & "            <pps:LastReplyDate>" & xSheet.Cells(StartRow, 27) & "</pps:LastReplyDate>"
            sEnv = sEnv & "            <pps:FacsimileTransmissionNumber>" & xSheet.Cells(StartRow, 28) & "</pps:FacsimileTransmissionNumber>"
            sEnv = sEnv & "            <pps:ProjectNumber>" & xSheet.Cells(StartRow, 29) & "</pps:ProjectNumber>"
            sEnv = sEnv & "            <pps:ProjectElement>" & xSheet.Cells(StartRow, 30) & "</pps:ProjectElement>"
            sEnv = sEnv & "            <pps:CurrencyTerms>" & xSheet.Cells(StartRow, 31) & "</pps:CurrencyTerms>"
            sEnv = sEnv & "            <pps:AgreedRate>" & xSheet.Cells(StartRow, 32) & "</pps:AgreedRate>"
            sEnv = sEnv & "            <pps:FutureRateAgreementNumber>" & xSheet.Cells(StartRow, 33) & "</pps:FutureRateAgreementNumber>"
            sEnv = sEnv & "            <pps:Agent>" & xSheet.Cells(StartRow, 34) & "</pps:Agent>"
            sEnv = sEnv & "            <pps:Payee>" & xSheet.Cells(StartRow, 35) & "</pps:Payee>"
            sEnv = sEnv & "            <pps:TermsText>" & xSheet.Cells(StartRow, 36) & "</pps:TermsText>"
            sEnv = sEnv & "            <pps:Signature>" & xSheet.Cells(StartRow, 37) & "</pps:Signature>"
            sEnv = sEnv & "            <pps:OrderTotalDiscountGenerating>" & xSheet.Cells(StartRow, 38) & "</pps:OrderTotalDiscountGenerating>"
            sEnv = sEnv & "            <pps:TotalOrderCost>" & xSheet.Cells(StartRow, 39) & "</pps:TotalOrderCost>"
            sEnv = sEnv & "            <pps:OrderTotalDiscount>" & xSheet.Cells(StartRow, 40) & "</pps:OrderTotalDiscount>"
            sEnv = sEnv & "         </pps:PPS200>"
            sEnv = sEnv & "      </pps:PPS200>"
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
'TODO: Mulesoft to return Purchase Order Number
'                    Set XmlNodeList = XmlDoc.getElementsByTagName("*")
'                    For Each XmlNode In XmlNodeList
'                        If XmlNode.nodeName = "PurchaseNumber" Then 'to check the right configuration to lock for in the respone
'                            If xSheet.Cells(StartRow, 5).Value = vbNullString Then
'                                xSheet.Cells(StartRow, 5).Value = XmlNode.Text
'                            End If
'                        End If
'                    Next XmlNode
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
    MsgBox "Process completed!", vbInformation, strMethod & " " & strTransaction
Else
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
