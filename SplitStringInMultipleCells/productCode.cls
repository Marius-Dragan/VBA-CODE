VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "productCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Rename this Class Module  productCode
'
'
Private pStyle As String
Private pFabric As String
Private pColour As String
Private pSize As String

Public Property Get Style() As String
    Style = pStyle
End Property
Public Property Let Style(value As String)
    pStyle = value
End Property

Public Property Get Fabric() As String
    Fabric = pFabric
End Property
Public Property Let Fabric(value As String)
    pFabric = UCase(value)
End Property

Public Property Get Colour() As String
    Colour = pColour
End Property
Public Property Let Colour(value As String)
    pColour = value
End Property

Public Property Get Size() As String
    Size = pSize
End Property
Public Property Let Size(value As String)
    pSize = value
End Property
