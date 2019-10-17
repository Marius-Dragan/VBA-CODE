Attribute VB_Name = "Module1"
Dim UpdateUserProfileError As Boolean
Dim responseError As String
Sub UploadHeaderMMS100()

Call UpdateUserProfile
If UpdateUserProfileError = False Then

    Dim responseText As String
    Dim sURL As String
    Dim sEnv As String
    Dim xmlhtp As New MSXML2.XMLHTTP60
    Dim xmlDoc As New DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim myNode As MSXML2.IXMLDOMNode

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
    
    Dim strTargetNamespace
    strTargetNamespace = strNamespaceBase & "/" & strServiceName & "/" & strMethod

    Dim strSoapAction
    strSoapAction = strWebServicesServer & "/" & strServiceName

    sURL = strServiceRoot & "/" & strServiceName
    
    Application.ScreenUpdating = False
    Dim StartRow As Long
    StartRow = Sheet2.Cells(5, 1)
    Dim EndRow As Long
    EndRow = Sheet2.Cells(7, 1)
    Do While (StartRow <= EndRow)
        ' Clear Return Code and Error Message
        Sheet2.Cells(StartRow, 2).ClearContents
        Sheet2.Cells(StartRow, 3).ClearContents
        
        'Add Controls to terminate the Loop when no information in the Cell!
        If IsEmpty(Sheet2.Cells(StartRow, 6)) = True Then
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
        sEnv = sEnv & "   <mms:MMS100>"
        sEnv = sEnv & "     <mms:MMS100>"
        
            If IsEmpty(Sheet2.Cells(StartRow, 5)) = False Then
                sEnv = sEnv & "       <mms:OrderNumber>" & Sheet2.Cells(StartRow, 5) & "</mms:OrderNumber>"
                Else
                sEnv = sEnv & "       <mms:OrderNumber>" & "" & "</mms:OrderNumber>"
            End If
            
        sEnv = sEnv & "      <mms:OrderType>" & Sheet2.Cells(StartRow, 6) & "</mms:OrderType>"
        sEnv = sEnv & "      <mms:Facility>" & Sheet2.Cells(StartRow, 7) & "</mms:Facility>"
        sEnv = sEnv & "      <mms:ToWarehouse>" & Sheet2.Cells(StartRow, 8) & "</mms:ToWarehouse>"
               
        If IsEmpty(Sheet2.Cells(StartRow, 9)) = False Then
            sEnv = sEnv & "     <mms:Remark>" & Sheet2.Cells(StartRow, 9) & "</mms:Remark>"
        End If
            
        sEnv = sEnv & "      <mms:ProjectNumber>" & Sheet2.Cells(StartRow, 10) & "</mms:ProjectNumber>"
        
        If IsEmpty(Sheet2.Cells(StartRow, 11)) = False Then
            sEnv = sEnv & "     <mms:ProjectElement>" & Sheet2.Cells(StartRow, 11) & "</mms:ProjectElement>"
        End If
        
        
            If IsEmpty(Sheet2.Cells(StartRow, 12)) = False Then
                sEnv = sEnv & "     <mms:ToLocation>" & Sheet2.Cells(StartRow, 12) & "</mms:ToLocation>"
            End If
            
            If IsEmpty(Sheet2.Cells(StartRow, 13)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderCategory>" & Sheet2.Cells(StartRow, 13) & "</mms:ReferenceOrderCategory>"
            End If
            
             If IsEmpty(Sheet2.Cells(StartRow, 14)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderNumber>" & Sheet2.Cells(StartRow, 14) & "</mms:ReferenceOrderNumber>"
            End If
            
             If IsEmpty(Sheet2.Cells(StartRow, 15)) = False Then
                sEnv = sEnv & "     <mms:ReferenceOrderLine>" & Sheet2.Cells(StartRow, 15) & "</mms:ReferenceOrderLine>"
            End If
        sEnv = sEnv & "   </mms:MMS100>"
        sEnv = sEnv & "</mms:MMS100>"
        sEnv = sEnv & "</soapenv:Body>"
        sEnv = sEnv & "</soapenv:Envelope>"


            With xmlhtp
                
                .Open "POST", sURL, False, strUsername, strPassword
                .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
                .setRequestHeader "soapAction", strSoapAction
                .setRequestHeader "Authorization", "Basic " + Base64Encode(strUsername + ":" + strPassword)
                .send sEnv
                
                xmlDoc.LoadXML (.responseText)
                
               
                If xmlhtp.Status = 200 Then
    
                    ' Success (200)
                   'MsgBox "OK " & xmlhtp.Status
                    Sheet2.Cells(StartRow, 2).Value = "OK"
                    
                    
                Else
                    ' Failure (404, 500, ...)
                    'MsgBox "NOK " & xmlhtp.Status
                    Sheet2.Cells(StartRow, 2).Value = "NOK"
                    Set xmlNodeList = xmlDoc.getElementsByTagName("*")
                    For Each xmlNode In xmlNodeList
                        For Each myNode In xmlNode.ChildNodes
                            If myNode.NodeType = NODE_TEXT Then
                                If xmlNode.nodeName = "faultstring" Then
                                    'MsgBox xmlNode.nodeName & "=" & xmlNode.Text
                                    Sheet2.Cells(StartRow, 3).Value = xmlNode.Text
                                End If
                            End If
                        Next myNode
                    Next xmlNode
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
Sub UploadLineMMS101Soap()

Call UpdateUserProfile
If UpdateUserProfileError = False Then

    Dim responseText As String
    Dim sURL As String
    Dim sEnv As String
    Dim xmlhtp As New MSXML2.XMLHTTP60
    Dim xmlDoc As New DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim myNode As MSXML2.IXMLDOMNode

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
                
                xmlDoc.LoadXML (.responseText)
    
                If xmlhtp.Status = 200 Then
                    ' Success (200)
                    'MsgBox "OK " & xmlhtp.Status
                    Sheet3.Cells(StartRow, 2).Value = "OK"
                Else
                    ' Failure (404, 500, ...)
                    'MsgBox "NOK " & xmlhtp.Status
                    Sheet3.Cells(StartRow, 2).Value = "NOK"
                    Set xmlNodeList = xmlDoc.getElementsByTagName("*")
                    For Each xmlNode In xmlNodeList
                        For Each myNode In xmlNode.ChildNodes
                            If myNode.NodeType = NODE_TEXT Then
                                If xmlNode.nodeName = "faultstring" Then
                                    'MsgBox xmlNode.nodeName & "=" & xmlNode.Text
                                    Sheet3.Cells(StartRow, 3).Value = xmlNode.Text
                                End If
                            End If
                        Next myNode
                    Next xmlNode
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
Sub UploadLineMMS101API()

Call UpdateUserProfile
If UpdateUserProfileError = False Then

Dim responseText As String
Dim sURL As String
Dim sRESTCall As String
Dim xmlhtp As Object
Dim xmlDoc As New MSXML2.DOMDocument60
Set xmlhtp = CreateObject("MSXML2.XMLHTTP")
Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
Dim webserviceSOAPActionNameSpace

Dim xmlNode As MSXML2.IXMLDOMNode
Dim xmlNodeList As MSXML2.IXMLDOMNodeList
Dim myNode As MSXML2.IXMLDOMNode


Dim strUsername
strUsername = "INFORBC\" & Sheet1.Cells(1, 2)
Dim strPassword
strPassword = Sheet1.Cells(2, 2)

'Capture the starting Row
Dim StartRow As Long
StartRow = Sheet3.Cells(5, 1)
Dim EndRow As Long
EndRow = Sheet3.Cells(7, 1)

Application.ScreenUpdating = False
Do While (StartRow <= EndRow)
        ' Clear Return Code and Error Message
        Sheet3.Cells(StartRow, 2).ClearContents
        Sheet3.Cells(StartRow, 3).ClearContents

sURL = Sheet1.Cells(4, 2)
sURL = sURL & ":63906/m3api-rest/execute/MMS100MI/AddDOLine?"
sURL = sURL & "&TRNR=" & Sheet3.Cells(StartRow, 5)
sURL = sURL & "&ITNO=" & Sheet3.Cells(StartRow, 7)
sURL = sURL & "&TRQT=" & Sheet3.Cells(StartRow, 9)

If Sheet3.Cells(StartRow, 10).Value <> "" Then
sURL = sURL & "&WHLO=" & Sheet3.Cells(StartRow, 10)
End If

If Sheet3.Cells(StartRow, 11).Value <> "" Then
sURL = sURL & "&WHSL=" & Sheet3.Cells(StartRow, 11)
End If

If Sheet3.Cells(StartRow, 12).Value <> "" Then
sURL = sURL & "&TWSL=" & Sheet3.Cells(StartRow, 12)
End If

If Sheet3.Cells(StartRow, 8).Value <> "" Then
sURL = sURL & "&RSCD=" & Sheet3.Cells(StartRow, 8)
End If


With xmlhtp

  .Open "GET", sURL, False, strUsername, strPassword
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "Authorization", "Basic " + Base64Encode(strUsername + ":" + strPassword)
    .send
    
    
    xmlDoc.LoadXML (.responseText)
    
    If xmlhtp.Status = 200 Then
           
           Set xmlNodeList = xmlDoc.getElementsByTagName("*")
             For Each xmlNode In xmlNodeList
               For Each myNode In xmlNode.ChildNodes
                        If xmlNode.nodeName = "Message" Then
                            Sheet3.Cells(StartRow, 2).Value = "NOK"
                            Sheet3.Cells(StartRow, 3).Value = xmlNode.Text
                        Else
                            Sheet3.Cells(StartRow, 2).Value = "OK"
                       End If
               Next myNode
            Next xmlNode
        Else
            Sheet3.Cells(StartRow, 2).Value = "NOK"
    End If
End With

sURL = ""
StartRow = StartRow + 1

Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, "M3 Upload"
Else
    MsgBox "Please check M3 user profile settings and try again." & vbNewLine & "Error: " & responseError, vbInformation, "M3 Profile Settings Update Error"
End If
End Sub
Sub ClearHeaderlogs()
    With Sheets("MMS100")
        Range(.Cells(4, 2), .Cells(500, 2)).ClearContents
        Range(.Cells(4, 3), .Cells(500, 3)).ClearContents
    End With
End Sub
Sub ClearLinelogs()
    With Sheets("MMS101")
        Range(.Cells(4, 2), .Cells(500, 2)).ClearContents
        Range(.Cells(4, 3), .Cells(500, 3)).ClearContents
    End With
End Sub
Sub UpdateUserProfile()
Dim USID, CONO, DIVI, FACI, WHLO
Dim strUsername, strPassword As String
Dim apiDetails As String
Dim HostURL As String
Dim setupWS As Worksheet
Dim headerWS As Worksheet
Dim lineWS As Worksheet
Dim sURL As String
Dim HeaderStartRow As Long
Dim LineStartRow As Long

Dim responseText As String
Dim xmlhtp As Object
Dim xmlDoc As New MSXML2.DOMDocument60
Set xmlhtp = CreateObject("MSXML2.XMLHTTP")
Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

Dim xmlNode As MSXML2.IXMLDOMNode
Dim xmlNodeList As MSXML2.IXMLDOMNodeList
Dim myNode As MSXML2.IXMLDOMNode

Set setupWS = Sheet1
Set headerWS = Sheet2
Set lineWS = Sheet3

HostURL = setupWS.Range("B4").Value & ":63906"

headerWS.Range("A19").ClearContents
strUsername = "INFORBC\" & UCase(setupWS.Range("B1").Value)
strPassword = setupWS.Range("B2").Value
HeaderStartRow = headerWS.Range("A5")
LineStartRow = lineWS.Range("A5")

         USID = CStr(setupWS.Range("B1").Value)
         'If USID <> vbNullString Then
            apiDetails = "?USID=" + UCase(USID)
         'End If
         
         CONO = CStr(headerWS.Range("A11").Value)
         If CONO <> vbNullString Then
            apiDetails = apiDetails + "&CONO=" + UCase(CONO)
         Else
            MsgBox "Company missing. Please enter a company and try again!", vbInformation, "M3 Error"
            Exit Sub
         End If
         
         DIVI = CStr(headerWS.Range("A13").Value)
         If DIVI <> vbNullString Then
            apiDetails = apiDetails + "&DIVI=" + UCase(DIVI)
         Else
            apiDetails = apiDetails + "&DIVI=" + CStr(headerWS.Range("G" & HeaderStartRow).Value)
            headerWS.Range("A13").Value = CStr(headerWS.Range("G" & HeaderStartRow).Value)
         End If

         FACI = CStr(headerWS.Range("A15").Value)
         If FACI <> vbNullString Then
            apiDetails = apiDetails + "&FACI=" + UCase(FACI)
         Else
            apiDetails = apiDetails + "&FACI=" + CStr(headerWS.Range("G" & HeaderStartRow).Value)
            headerWS.Range("A15").Value = CStr(headerWS.Range("G" & HeaderStartRow).Value)
         End If
         
         WHLO = CStr(headerWS.Range("A17").Value)
         If WHLO <> vbNullString Then
            apiDetails = apiDetails + "&WHLO=" + UCase(WHLO)
         Else
            apiDetails = apiDetails + "&WHLO=" + CStr(lineWS.Range("J" & LineStartRow).Value)
            headerWS.Range("A17").Value = CStr(lineWS.Range("J" & LineStartRow).Value)
         End If

sURL = HostURL + "/m3api-rest/execute/MNS150MI/" & "ChgDefaultValue" & apiDetails

With xmlhtp

  .Open "GET", sURL, False, strUsername, strPassword
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "Authorization", "Basic " + Base64Encode(strUsername + ":" + strPassword)
    .send
    
    
    xmlDoc.LoadXML (.responseText)
    
    If xmlhtp.Status = 200 Then
           
           Set xmlNodeList = xmlDoc.getElementsByTagName("*")
             For Each xmlNode In xmlNodeList
               For Each myNode In xmlNode.ChildNodes
                        If xmlNode.nodeName = "Message" Then
                            WS.Range("A19").Value = "NOK"
                            WS.Range("A19").Value = xmlNode.Text
                            WS.Range("A19").Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
                            WS.Range("A19").Replace "   ", ""
                            responseError = WS.Range("A19")
                            UpdateUserProfileError = True
                        Else
                            headerWS.Range("A19").Value = "OK"
                            UpdateUserProfileError = False
                       End If
               Next myNode
            Next xmlNode
        Else
            headerWS.Range("A19").Value = "NOK"
    End If
End With
End Sub
Function Base64Encode(sText)
    Dim oXml, oNode
    Set oXml = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXml.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.Text
    Set oNode = Nothing
    Set oXml = Nothing
End Function
Function Stream_StringToBinary(Text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string

Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function

