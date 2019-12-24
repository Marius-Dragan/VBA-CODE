Attribute VB_Name = "MMS100MI_HEAD"
Option Explicit
Public UpdateUserProfileError As Boolean
Public responseError As String
Sub UploadHeaderMMS100()

Call UpdateUserProfile
If UpdateUserProfileError = False Then

    Dim sURL As String
    Dim sEnv As String
    Dim xSheet As Worksheet
    Dim strTransaction As String


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
    strServiceName = "DistributionOrders"
    Dim strMethod
    strMethod = "MMS100"

    Dim strTargetNamespace
    strTargetNamespace = strNamespaceBase & "/" & strServiceName & "/" & strMethod

    Dim strSoapAction
    strSoapAction = strWebServicesServer & "/" & strServiceName

    sURL = strServiceRoot & "/" & strServiceName

    Application.ScreenUpdating = False

    Do While (StartRow <= EndRow)

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
        sEnv = sEnv & "   <mms:MMS100>"
        sEnv = sEnv & "     <mms:MMS100>"

            If IsEmpty(xSheet.Cells(StartRow, 3)) = False Then
                sEnv = sEnv & "       <mms:OrderNumber>" & xSheet.Cells(StartRow, 3) & "</mms:OrderNumber>"
                Else
                sEnv = sEnv & "       <mms:OrderNumber>" & "" & "</mms:OrderNumber>"
            End If

        sEnv = sEnv & "      <mms:OrderType>" & xSheet.Cells(StartRow, 4) & "</mms:OrderType>"
        sEnv = sEnv & "      <mms:Facility>" & xSheet.Cells(StartRow, 5) & "</mms:Facility>"
        sEnv = sEnv & "      <mms:ToWarehouse>" & xSheet.Cells(StartRow, 6) & "</mms:ToWarehouse>"

        If IsEmpty(xSheet.Cells(StartRow, 7)) = False Then
            sEnv = sEnv & "     <mms:Remark>" & xSheet.Cells(StartRow, 7) & "</mms:Remark>"
        End If

        sEnv = sEnv & "      <mms:ProjectNumber>" & xSheet.Cells(StartRow, 8) & "</mms:ProjectNumber>"

        If IsEmpty(xSheet.Cells(StartRow, 9)) = False Then
            sEnv = sEnv & "     <mms:ProjectElement>" & xSheet.Cells(StartRow, 9) & "</mms:ProjectElement>"
        End If


            If IsEmpty(xSheet.Cells(StartRow, 10)) = False Then
                sEnv = sEnv & "     <mms:ToLocation>" & xSheet.Cells(StartRow, 10) & "</mms:ToLocation>"
            End If

            If IsEmpty(xSheet.Cells(StartRow, 11)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderCategory>" & xSheet.Cells(StartRow, 11) & "</mms:ReferenceOrderCategory>"
            End If

             If IsEmpty(xSheet.Cells(StartRow, 12)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderNumber>" & xSheet.Cells(StartRow, 12) & "</mms:ReferenceOrderNumber>"
            End If

             If IsEmpty(xSheet.Cells(StartRow, 13)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderLine>" & xSheet.Cells(StartRow, 13) & "</mms:ReferenceOrderLine>"
            End If
        sEnv = sEnv & "   </mms:MMS100>"
        sEnv = sEnv & "</mms:MMS100>"
        sEnv = sEnv & "</soapenv:Body>"
        sEnv = sEnv & "</soapenv:Envelope>"


            With XmlHttp

                .Open "POST", sURL, False, strUsername, strPassword
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store chache
                .setRequestHeader "soapAction", strSoapAction
                .setRequestHeader "Authorization", "Basic " + Encoding.Base64Encode(strUsername + ":" + strPassword)
                .send sEnv

                XmlDoc.LoadXML (.responseText)

                If XmlHttp.Status = 200 Then ' Success (200)
                    xSheet.Cells(StartRow, 1).Value = "OK"
                    Set XmlNodeList = XmlDoc.getElementsByTagName("*")
                    For Each XmlNode In XmlNodeList
                        If XmlNode.nodeName = "OrderNumber" Then
                            If xSheet.Cells(StartRow, 3).Value = vbNullString Then
                                xSheet.Cells(StartRow, 3).NumberFormat = "@"
                                xSheet.Cells(StartRow, 3).Value = XmlNode.Text
                            End If
                        End If
                    Next XmlNode
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
