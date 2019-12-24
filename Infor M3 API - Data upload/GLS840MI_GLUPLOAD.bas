Attribute VB_Name = "GLS840MI_GLUPLOAD"
Private Type invdetail
       GLAcct As String
       CostCentre As String
       Channel As String
       ProdCateg As String
       Season As String
       InterCo As String
       Dim7 As String
       Desc As String
       Debit  As Double
       Credit As Double
End Type



Dim ProcessDate As String
Dim ProcessTime As String
Dim HostURL As String
Dim gnnr As String
Dim company As String
Dim APID As String
Dim GLDate As Date
Dim division As String
Dim Curr As String
Dim Chart As String
Dim Book As String
Dim Database As String
Dim Category As String
Dim BalanceType As String
Dim JournalName As String
Dim DataAccess As String
Dim Source As String
Dim LoginId As String
Dim UserId As String
Dim Password As String

Dim reqNO, lineno, Key, HeaderDesc, interface
Dim RNNO, GRNR, DIVI, SUNO, SPYN, SINO, IVDT, IVAM, VTTX, VTAM, CUAM, CUCD, CRTP, ACDT, AIT1, AIT2, AIT3, AIT4, AIT5, AIT6, AIT7, ARAT, APCD
Dim Div, Fac, GLAcct
Dim TotalAmt, CostCentre, Error
Dim ivd(1000) As invdetail, ix As Integer, Z As Integer
Sub GLS840_GLUPLOAD()
'
' upload
'

    Dim myFile As String, rng As Range, cellValue As Variant, i As Integer, j As Integer
    Dim LRow As Long
    Dim xSheet As Worksheet
    Dim versionNo As String

    Set xSheet = Sheet1
    versionNo = "GL-V2.4"

    LRow = xSheet.Range("B" & xSheet.Rows.Count).End(xlUp).Row

    HostURL = "https://yourdomaindev.cloud.infor.com:12345"
    interface = "GLUPLOAD"
    'PD = DateValue(Range("I5").Cells.Value)
    CD = Date
    Range("I6").Cells.Value = ""
    CT = Time()
    GRNR = Format(Month(CD), "00") + Format(Day(CD), "00") + Format(Hour(CT), "00") + Format(Minute(CT), "00")

    ProcessDate = Right(Format(Year(CD), "0000"), 2) + Format(Month(CD), "00") + Format(Day(CD), "00")
    ProcessTime = Format(Hour(CT), "00") + Format(Minute(CT), "00")

    company = "100"

    Range("H8").Value = "Version No"
    Range("I8").Value = versionNo
    division = Range("C4").Cells.Value
    GLDate = Range("F4").Cells.Value
    Curr = Range("F6").Cells.Value
    'Chart = Range("D8").Cells.Value
    'Book = Range("D9").Cells.Value
    'Database = Range("D10").Cells.Value
    'Category = Range("D11").Cells.Value
    APID = Range("C6").Cells.Value
      If IsNumeric(APID) Then
          APID = Format(APID, "00000")
      End If
    UserId = UCase(APID)
    LoginId = "INFORBC\" & UserId
    'BalanceType = Range("F6").Cells.Value
    JournalName = Range("C8").Cells.Value
    'DataAccess = Range("G8").Cells.Value
    'Source = Range("G9").Cells.Value

    Password = Application.InputBox("Enter a Password:")


    RNNO = "0" + Format(Month(CD), "00") + Format(Day(CD), "00") + Format(Hour(CT), "00") + Format(Minute(CT), "00")
    lineno = 0
    ix = 0
    Key = Left(UserId, 5) + ProcessDate + ProcessTime
    HeaderDesc = JournalName


     request = HostURL + "/m3api-rest/execute/GLS840MI/AddBatchHead?CONO=" + company & _
             "&DIVI=" + division & _
             "&KEY1=" + Key & _
             "&INTN=" + interface & _
             "&DESC=" + HeaderDesc & _
             "&USID=" + UserId

     Webservice (request)
     If Error <> True Then

'Old way of looping not efficient to track
'     For Each rng In Range("B12:K1012").Rows
'      If rng.Cells(i, 1).Value Then
'       ix = ix + 1
'        'i = ix
'        Debug.Print ix
'        Debug.Print i
'       ivd(ix).GLAcct = rng.Cells(i, 1).Value        'B
'       ivd(ix).Facility = rng.Cells(i, 2).Value      'C
'       ivd(ix).InterCo = rng.Cells(i, 3).Value       'D
'       ivd(ix).ProdCateg = rng.Cells(i, 4).Value       'E
'       ivd(ix).Season = rng.Cells(i, 5).Value           'F
'       ivd(ix).Channel = rng.Cells(i, 6).Value       'G
'       ivd(ix).Dim7 = rng.Cells(i, 7).Value          'H
'       ivd(ix).Desc = rng.Cells(i, 8).Value          'I
'       ivd(ix).Debit = rng.Cells(i, 9).Value         'J
'       ivd(ix).Credit = rng.Cells(i, 10).Value       'K
'
'     End If
'
'     Next

    For i = 12 To LRow
        With xSheet
            If .Cells(i, 2).Value <> vbNullString Then
                ix = ix + 1

                ivd(ix).GLAcct = .Range("B" & i).Value
                ivd(ix).CostCentre = .Range("C" & i).Value
                ivd(ix).Channel = .Range("D" & i).Value
                ivd(ix).ProdCateg = .Range("E" & i).Value
                ivd(ix).Season = .Range("F" & i).Value
                ivd(ix).InterCo = .Range("G" & i).Value
                ivd(ix).Dim7 = .Range("H" & i).Value
                ivd(ix).Desc = .Range("I" & i).Value
                ivd(ix).Debit = .Range("J" & i).Value
                ivd(ix).Credit = .Range("K" & i).Value
            End If
        End With
    Next i


     writerecords

      Range("I6").Cells.Value = CD + CT
    End If
    MsgBox "Process completed!", vbInformation
End Sub

Public Sub writerecords()
    For Z = 1 To ix
        FormatI1
    Next
End Sub


Public Sub FormatI1()
   lineno = lineno + 1
       TotalAmt = ivd(Z).Debit - ivd(Z).Credit


       DIVI = division
       CUCD = Curr
       CRTP = "01"
       If Sheet1.Range("K8").Value = "dot" Then
            CUAM = Rept(" ", 17 - Len(Format(TotalAmt, "0.00"))) + Format(TotalAmt, "0.00")
       Else
            CUAM = Rept(" ", 17 - Len(Format(TotalAmt, "0.00"))) + Replace(Format(TotalAmt, "0.00"), ".", ",")
       End If
       ACDT = Format(Year(GLDate), "0000") + Format(Month(GLDate), "00") + Format(Day(GLDate), "00")
       AIT1 = Left(Format(ivd(Z).GLAcct) + "          ", 10)
       AIT2 = Left(ivd(Z).CostCentre + "          ", 10)
       AIT3 = Left(ivd(Z).Channel + "          ", 10)
       AIT4 = Left(ivd(Z).ProdCateg + "          ", 10)
       AIT5 = Left(ivd(Z).Season + "          ", 10)
       AIT6 = Left(ivd(Z).InterCo + "          ", 10)
       AIT7 = Left(ivd(Z).Dim7 + "          ", 10)
       VTXT = Left(ivd(Z).Desc + Rept(" ", 40), 40)

      parm = "I1" + RNNO + GRNR + DIVI + _
       AIT1 + AIT2 + AIT3 + AIT4 + AIT5 + AIT6 + AIT7 + CUCD + CUAM + ACDT + VTXT

            request = HostURL + "/m3api-rest/execute/GLS840MI/AddBatchLine?CONO=" + company & _
             "&DIVI=" + division & _
             "&KEY1=" + Key & _
             "&LINE=" + Format(lineno) & _
             "&PARM=" + parm
        Webservice (request)
End Sub


Public Sub Webservice(request)
 'MsgBox "REQUEST:" + request
 Error = False

'define XML and HTTP components

Dim M3X As Object
Dim M3Service As Object

    'Dim XmlDoc As Object
    'Dim HttpClient As Object
    Set M3X = CreateObject("MSXML2.DOMDocument.6.0")
    Set M3Service = CreateObject("MSXML2.XMLHTTP.6.0")

'create HTTP request to query URL - make sure to have
'that last "False" there for synchronous operation

M3Service.Open "GET", request, False, LoginId, Password
M3Service.setRequestHeader "Content-Type", "application/xml"
M3Service.setRequestHeader "Cache-Control", "no-cache" 'Force IE not to store cache
M3Service.setRequestHeader "Authorization", "Basic " + Base64Encode(LoginId + ":" + Password)

'send HTTP request

M3Service.send
 Reply = M3Service.responseText
 'MsgBox "REPLY:" + Reply + ":END"

M3X.LoadXML Reply
ReplyName = M3X.DocumentElement.nodeName
If ReplyName = "ErrorMessage" Then

  MSG = M3X.DocumentElement.FirstChild.Text
  MsgBox MSG
  Error = True
End If



End Sub
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.Text
    Set oNode = Nothing
    Set oXML = Nothing
End Function


'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
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
Public Function Rept(str, x) As String
   Rept = Application.WorksheetFunction.Rept(str, x)
End Function
