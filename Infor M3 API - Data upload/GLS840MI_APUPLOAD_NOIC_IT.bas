Attribute VB_Name = "GLS840MI_APUPLOAD_NOIC_IT"
Private Type invdetail
       LastInvoiceNo As String
       InvDesc As String
       invDate As Date
       dueDate As Date 'NEW
       'GLDate As Date  'CHANGED
       Curr As String 'new
       ExchRate As Double   'new
       InvoiceAmt As Double
       Auth As String
       GLAcct As String
       Dim2 As String
       Dim4 As String
       Dim5 As String
       Dim6 As String
       VATcode As String
End Type

Dim ProcessDate As String
Dim ProcessTime As String
Dim HostURL As String
Dim gnnr As String
Dim company As String
Dim APID
Dim division, UserId, LoginId, Password, lineno, Key, HeaderDesc, interface
Dim RNNO, GRNR, DIVI, SUNO, SPYN, SINO, IVDT, DUDT, IVAM, VTCD, VTAM, CUCD, CRTP, ACDT, AIT1, AIT2, AIT4, AIT5, AIT6, ARAT, APCD
Dim InvDesc, invDate, dueDate, GLDate, InvoiceAmt, Auth, Curr, ExchRate, Div, GLAcct, Dim2, Dim4, Dim5, Dim6, VATcode  'removed Facility
Dim LastInvoiceNo, LastSupplier, supplier, TotalAmt, InvoiceNo, Error
Dim ivd(3500) As invdetail, ix As Integer, Z As Integer, SeqNo As Long
Sub SendToM3()
'
' upload
'

'
    Dim myFile As String, rng As Range, cellValue As Variant, i As Integer, j As Integer

    HostURL = "https://yourdomain.cloud.infor.com:12345"    'Prod Env
    'hosturl = "https://yourdomaindev.cloud.infor.com:12345"  'Dev Env
    interface = "APUPLD-NOIC-IT"
    PD = DateValue(Range("F4").Cells.Value)
    CD = Date
    Range("F6").Cells.Value = ""

    CT = Time()
    'GRNR = Format(Month(CD), "00") + Format(Day(CD), "00") + Format(Hour(CT), "00") + Format(Minute(CT), "00") 'added to the lines to generate different voucher number per invoice

    ProcessDate = Right(Format(Year(CD), "0000"), 2) + Format(Month(CD), "00") + Format(Day(CD), "00")
    ProcessTime = Format(Hour(CT), "00") + Format(Minute(CT), "00")
    company = "100"
    APID = Range("I6").Cells.Value 'username
    'Facility = Range("I4").Cells.Value  'commented
    division = Range("C4").Cells.Value
    GLDate = Range("C6").Cells.Value   'new
    'Curr = Range("C8").Cells.Value 'commented
    If IsNumeric(APID) Then
        APID = Format(APID, "00000")
    End If
    UserId = APID
    LoginId = "INFORBC\" & UserId

    Password = Application.InputBox("Enter a Password:")

    'RNNO = "0000" + Format(APID, "00000")
    RNNO = "0" + Format(Month(CD), "00") + Format(Day(CD), "00") + Format(Hour(CT), "00") + Format(Minute(CT), "00")

    lineno = 0
    ix = 0
    SeqNo = 0

    Key = Left(UserId, 5) + ProcessDate + ProcessTime
    HeaderDesc = "Invoice upload " + ProcessDate + UserId


     request = HostURL + "/m3api-rest/execute/GLS840MI/AddBatchHead?CONO=" + company & _
             "&DIVI=" + division & _
             "&KEY1=" + Key & _
             "&INTN=" + interface & _
             "&DESC=" + HeaderDesc & _
             "&USID=" + UserId

     Webservice (request)
     If Error <> True Then

     For Each rng In Range("B10:P3015").Rows    'changed to column L, and B10

     supplier = rng.Cells(i, 1).Value  'B
     InvoiceNo = rng.Cells(i, 2) 'C
     'Debug.Print InvoiceNo

     If supplier <> "" Then
      If supplier <> LastSupplier Or InvoiceNo <> LastInvoiceNo Then
        If LastInvoiceNo <> "" Then
          FormatI1
          writerecords
          ix = 0
        End If

         LastInvoiceNo = InvoiceNo
         'Debug.Print SeqNo
         SeqNo = SeqNo + 1
         LastSupplier = supplier
         TotalAmt = 0
     End If
     ix = ix + 1

       ivd(ix).InvDesc = rng.Cells(i, 3).Value    'D
       ivd(ix).invDate = DateValue(rng.Cells(i, 4).Value)    'E
       ivd(ix).dueDate = DateValue(rng.Cells(i, 5).Value)    'F
       'ivd(ix).GLDate = DateValue(rng.Cells(i, 6).Value)     'G
       ivd(ix).Curr = rng.Cells(i, 6).Value  'G  'new
       ivd(ix).ExchRate = rng.Cells(i, 7).Value 'H 'new
       ivd(ix).InvoiceAmt = rng.Cells(i, 8).Value 'I
       ivd(ix).Auth = rng.Cells(i, 9).Value  'j
       ivd(ix).GLAcct = rng.Cells(i, 10).Value    'K
       ivd(ix).Dim2 = rng.Cells(i, 11).Value    'L
       ivd(ix).Dim4 = rng.Cells(i, 12).Value    'M
       ivd(ix).Dim5 = rng.Cells(i, 13).Value    'N
       ivd(ix).Dim6 = rng.Cells(i, 14).Value    'O
       ivd(ix).VATcode = rng.Cells(i, 15).Value    'P

       InvDesc = rng.Cells(i, 3).Value    'D
       invDate = DateValue(rng.Cells(i, 4).Value)    'E
       dueDate = DateValue(rng.Cells(i, 5).Value)    'F
       'GLDate = DateValue(rng.Cells(i, 6).Value)     'G
       Curr = rng.Cells(i, 6).Value  'G 'new
       ExchRate = rng.Cells(i, 7).Value 'H 'new
       InvoiceAmt = rng.Cells(i, 8).Value 'I
       Auth = rng.Cells(i, 9).Value  'J
       GLAcct = rng.Cells(i, 10).Value    'K

       Dim2 = rng.Cells(i, 11).Value    'L
       Dim4 = rng.Cells(i, 12).Value    'M
       Dim5 = rng.Cells(i, 13).Value    'N
       Dim6 = rng.Cells(i, 14).Value    'O
       VATcode = rng.Cells(i, 15).Value    'P

       ' FormatI2
       TotalAmt = TotalAmt - ivd(ix).InvoiceAmt
     End If
     Next

    If LastInvoiceNo <> "" Then
       FormatI1
       writerecords
    End If
     Range("F6").Cells.Value = CD + CT
     MsgBox "Upload Completed!", vbOKOnly, "M3"
    End If
    LastInvoiceNo = ""
End Sub

Public Sub writerecords()
    For Z = 1 To ix
        FormatI2
    Next
End Sub


Public Sub FormatI1()
   lineno = lineno + 1

       DIVI = division
       GRNR = Left(Format(SeqNo) + "        ", 8)
       SUNO = Left(LastSupplier & "          ", 10)
       'SUNO = LastSupplier + Rept(" ", 10 - Len(LastSupplier))
       SPYN = SUNO
       'SINO = LastInvoiceNo + Rept(" ", 24 - Len(LastInvoiceNo))
       SINO = Left(LastInvoiceNo + "                        ", 24)
       IVDT = Format(Year(invDate), "0000") + Format(Month(invDate), "00") + Format(Day(invDate), "00")
       DUDT = Format(Year(dueDate), "0000") + Format(Month(dueDate), "00") + Format(Day(dueDate), "00")
       IVAM = Rept(" ", 17 - Len(Format(-TotalAmt, "0.00"))) + Format(-TotalAmt, "0.00")
       VTCD = Left(Format(VATcode) + "  ", 2)
       VTAM = "                0"
       CUCD = Curr
       CRTP = "01"
       ARAT = Rept(" ", 11 - Len(Format(ExchRate, "0.000000"))) + Format(ExchRate, "0.000000")  'Changed from "     1"
       APCD = Left(Format(Auth) + "          ", 10)
       ACDT = Format(Year(GLDate), "0000") + Format(Month(GLDate), "00") + Format(Day(GLDate), "00")
       'AIT1 = Left(Format(GLAcct) + "          ", 10)   'new
       AIT1 = "21010     "   'if AP account is hard coded
       'AIT4 = Left(Format(Dim4) + "          ", 10)
       'AIT5 = Left(Format(Dim5) + "          ", 10)
       'AIT6 = Left(Format(Dim6) + "          ", 10)
       AIT6 = "          "   'if AP account is hard coded
       'AIT7 = Left(Format(Dim7) + "          ", 10)

      parm = "I1" + RNNO + GRNR + DIVI + _
       SUNO + SPYN + SINO + IVDT + DUDT + IVAM + VTCD + VTAM + CUCD + ARAT + ACDT + APCD + AIT1 + AIT6

            request = HostURL + "/m3api-rest/execute/GLS840MI/AddBatchLine?CONO=" + company & _
             "&DIVI=" + division & _
             "&KEY1=" + Key & _
             "&LINE=" + Format(lineno) & _
             "&PARM=" + parm
        Webservice (request)
End Sub

Public Sub FormatI2()
     lineno = lineno + 1
       DIVI = division
       GRNR = Left(Format(SeqNo) + "        ", 8)
       SUNO = Left(LastSupplier & "          ", 10)
       'SUNO = LastSupplier + Rept(" ", 10 - Len(LastSupplier))
       SPYN = SUNO
       'SINO = InvoiceNo + Rept(" ", 24 - Len(InvoiceNo))
       SINO = Left(Format(InvoiceNo) + "                        ", 24)
       IVDT = Format(Year(ivd(Z).invDate), "0000") + Format(Month(ivd(Z).invDate), "00") + Format(Day(ivd(Z).invDate), "00")
       DUDT = Format(Year(ivd(Z).dueDate), "0000") + Format(Month(ivd(Z).dueDate), "00") + Format(Day(ivd(Z).dueDate), "00")
       IVAM = Rept(" ", 17 - Len(Format(ivd(Z).InvoiceAmt, "0.00"))) + Format(ivd(Z).InvoiceAmt, "0.00")
       'VTCD = " 0"
       VTCD = Left(Format(ivd(Z).VATcode) + "  ", 2)
       VTAM = "                0"
       CUCD = Curr
       CRTP = "01"
       ARAT = Rept(" ", 11 - Len(Format(ExchRate, "0.000000"))) + Format(ExchRate, "0.000000")  'Changed from "     1"
       APCD = "          "
       ACDT = Format(Year(GLDate), "0000") + Format(Month(GLDate), "00") + Format(Day(GLDate), "00")
       'ACDT = Format(Year(ivd(Z).GLDate), "0000") + Format(Month(ivd(Z).GLDate), "00") + Format(Day(ivd(Z).GLDate), "00")
       AIT1 = Left(Format(ivd(Z).GLAcct) + "        ", 10) 'changed
       'AIT1 = "99997     "
       AIT2 = Left(Format(ivd(Z).Dim2) + "          ", 10)
       AIT4 = Left(Format(ivd(Z).Dim4) + "          ", 10)
       AIT5 = Left(Format(ivd(Z).Dim5) + "          ", 10)
       AIT6 = Left(Format(ivd(Z).Dim6) + "          ", 10)
       InvDesc = Format(ivd(Z).InvDesc)

      parm = "I2" + RNNO + GRNR + DIVI + _
       SUNO + SPYN + "                        " + IVDT + DUDT + IVAM + VTCD + VTAM + CUCD + ARAT + ACDT + APCD + AIT1 + AIT2 + AIT4 + AIT5 + AIT6 + InvDesc

      request = HostURL + "/m3api-rest/execute/GLS840MI/AddBatchLine?CONO=" + company & _
      "&DIVI=" + division & _
      "&KEY1=" + Key & _
      "&LINE=" + Format(lineno) & _
      "&PARM=" + parm
       Webservice (request)
End Sub



Public Sub Webservice(request)
 'MsgBox request
 Error = False

'define XML and HTTP components

Dim M3X As New DOMDocument60
Dim M3Service As New XMLHTTP60

'create HTTP request to query URL - make sure to have
'that last "False" there for synchronous operation
With M3Service
    .Open "GET", request, False, LoginId, Password
    .setRequestHeader "Content-Type", "application/xml"
    .setRequestHeader "Authorization", "Basic " + Base64Encode(LoginId + ":" + Password)
    .send
     Reply = .responseText
End With
'send HTTP request


'  MsgBox Reply
With M3X
    .LoadXML Reply
    ReplyName = .DocumentElement.nodeName
    If ReplyName = "ErrorMessage" Then
      MSG = .DocumentElement.FirstChild.Text
      MsgBox MSG
      Error = True
    End If
End With

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
