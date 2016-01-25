VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductieSoort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private Pnaam As String
Private pKleur As Long

'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get Id() As Long
    Id = pId
End Property

Public Property Let Id(Value As Long)
    pId = Value
End Property

'''''''''''''''''''''''''
' Naam property
'''''''''''''''''''''''''

Public Property Get Naam() As String
    Naam = Pnaam
End Property

Public Property Let Naam(Value As String)
    Pnaam = Value
End Property

'''''''''''''''''''''''''
' Kleur property
'''''''''''''''''''''''''

Public Property Get Kleur() As Long
    Kleur = pKleur
End Property

Public Property Let Kleur(Value As Long)
    pKleur = Value
End Property

