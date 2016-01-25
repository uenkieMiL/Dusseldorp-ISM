VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vestiging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As String
Private pOmschrijving As String

'''''''''''''''''''''''''
' Synergy property
'''''''''''''''''''''''''

Public Property Get Id() As String
    Id = pId
End Property

Public Property Let Id(Value As String)
    pId = Value
End Property

'''''''''''''''''''''''''
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property
