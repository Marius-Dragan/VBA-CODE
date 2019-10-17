Attribute VB_Name = "UpdatePI"
Sub UpdateMMS310MI()

Dim sURL As String
Dim Reply, ReplyName, MSG As String
Set M3X = CreateObject("MSXML2.DOMDocument.6.0")
Set M3Service = CreateObject("MSXML2.XMLHTTP.6.0")

Dim strUsername
strUsername = "INFORBC\" & Sheet1.Cells(1, 9)
Dim strPassword
strPassword = Sheet1.Cells(2, 9)

'Capture the starting Row
Dim StartRow As Long
StartRow = Sheet1.Cells(1, 2)
Dim EndRow As Long
EndRow = Sheet1.Cells(2, 2)

Application.ScreenUpdating = False
Do While (StartRow <= EndRow)

    If Sheet1.Range("L2").Value = "Production" Then
        sURL = "https://stellamc-bel1.cloud.infor.com:63906/m3api-rest/execute/MMS310MI/Update?"
    Else
        sURL = "https://stellamcdev-bel1.cloud.infor.com:63906/m3api-rest/execute/MMS310MI/Update?"
    End If

sURL = sURL & "&CONO=" & Sheet1.Cells(StartRow, 3)
sURL = sURL & "&WHLO=" & Sheet1.Cells(StartRow, 4)
sURL = sURL & "&ITNO=" & Sheet1.Cells(StartRow, 5)

If Sheet1.Cells(StartRow, 6).Value <> vbNullString Then
sURL = sURL & "&WHSL=" & Sheet1.Cells(StartRow, 6)
End If

If Sheet1.Cells(StartRow, 7).Value <> vbNullString Then
sURL = sURL & "&BANO=" & Sheet1.Cells(StartRow, 7)
End If

If Sheet1.Cells(StartRow, 8).Value <> vbNullString Then
sURL = sURL & "&CAMU=" & Sheet1.Cells(StartRow, 8)
End If

If Sheet1.Cells(StartRow, 9).Value <> vbNullString Then
sURL = sURL & "&REPN=" & Sheet1.Cells(StartRow, 9)
End If

If Sheet1.Cells(StartRow, 10).Value <> vbNullString Then
sURL = sURL & "&STQI=" & Sheet1.Cells(StartRow, 10)
End If

If Sheet1.Cells(StartRow, 11).Value <> vbNullString Then
sURL = sURL & "&STAG=" & Sheet1.Cells(StartRow, 11)
End If

If Sheet1.Cells(StartRow, 12).Value <> vbNullString Then
sURL = sURL & "&CAWI=" & Sheet1.Cells(StartRow, 12)
End If

If Sheet1.Cells(StartRow, 13).Value <> vbNullString Then
sURL = sURL & "&STDI=" & Sheet1.Cells(StartRow, 13)
End If

If Sheet1.Cells(StartRow, 14).Value <> vbNullString Then
sURL = sURL & "&TIHH=" & Sheet1.Cells(StartRow, 14)
End If

If Sheet1.Cells(StartRow, 15).Value <> vbNullString Then
sURL = sURL & "&TIMM=" & Sheet1.Cells(StartRow, 15)
End If

If Sheet1.Cells(StartRow, 16).Value <> vbNullString Then
sURL = sURL & "&TISS=" & Sheet1.Cells(StartRow, 16)
End If

If Sheet1.Cells(StartRow, 17).Value <> vbNullString Then
sURL = sURL & "&PRDT=" & Sheet1.Cells(StartRow, 17)
End If

If Sheet1.Cells(StartRow, 18).Value <> vbNullString Then
sURL = sURL & "&TRPR=" & Sheet1.Cells(StartRow, 18)
End If

If Sheet1.Cells(StartRow, 19).Value <> vbNullString Then
sURL = sURL & "&BREF=" & Sheet1.Cells(StartRow, 19)
End If

If Sheet1.Cells(StartRow, 20).Value <> vbNullString Then
sURL = sURL & "&BRE2=" & Sheet1.Cells(StartRow, 20)
End If

If Sheet1.Cells(StartRow, 21).Value <> vbNullString Then
sURL = sURL & "&BREM=" & Sheet1.Cells(StartRow, 21)
End If

If Sheet1.Cells(StartRow, 22).Value <> vbNullString Then
sURL = sURL & "&RSCD=" & Sheet1.Cells(StartRow, 22)
End If


With M3Service
    .Open "GET", sURL, False, strUsername, strPassword
    .SetRequestHeader "Content-Type", "application/xml"
    .SetRequestHeader "Authorization", "Basic " + Base64Encode(strUsername + ":" + strPassword)
    .send 'send HTTP request
    Reply = .responseText
End With
 
'Debug.Print Reply
With M3X
    .LoadXML Reply
    ReplyName = .DocumentElement.nodeName
    If ReplyName = "ErrorMessage" Then
      
        MSG = .DocumentElement.FirstChild.Text
      'Debug.Print msg
      Sheet1.Range("B" & StartRow).Value = MSG
      Sheet1.Range("B" & StartRow).Replace What:=Chr(160), Replacement:=Chr(32), LookAt:=xlPart, SearchOrder:=xlByColumns, MatchCase:=False
      Sheet1.Range("B" & StartRow).Replace "  ", ""
      Sheet1.Range("A" & StartRow).Value = "NOK"
    Else
        MSG = .DocumentElement.FirstChild.Text
        Sheet1.Range("B" & StartRow).Value = MSG & " Uploaded OK"
        Sheet1.Range("A" & StartRow).Value = "OK"
    End If
End With

sURL = ""
StartRow = StartRow + 1

Loop
    Application.ScreenUpdating = True
    MsgBox "Process completed!", vbInformation, "UpdateMMS310MI"
End Sub
Sub ClearHeaderlogs()
    With Sheet1
        .Range(.Cells(6, 1), .Cells(5000, 2)).ClearContents
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



