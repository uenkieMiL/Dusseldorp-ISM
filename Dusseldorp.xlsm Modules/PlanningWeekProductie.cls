VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanningWeekProductie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pKolomStart As Long
Private pKolomEind As Long
Private pStartDatum As Date
Private pEindDatum As Date
Private pKleur As Long
Private pSoort As String

''''''''''''''''''''''
' KolomStart property
''''''''''''''''''''''
Public Property Get KolomStart() As Long
    KolomStart = pKolomStart
End Property
Public Property Let KolomStart(Value As Long)
    pKolomStart = Value
End Property

''''''''''''''''''''''
' KolomEind property
''''''''''''''''''''''
Public Property Get KolomEind() As Long
    KolomEind = pKolomEind
End Property
Public Property Let KolomEind(Value As Long)
    pKolomEind = Value
End Property

''''''''''''''''''''''
' Kleur property
''''''''''''''''''''''
Public Property Get Kleur() As Long
    Kleur = pKleur
End Property
Public Property Let Kleur(Value As Long)
    pKleur = Value
End Property

''''''''''''''''''''''
' Soort property
''''''''''''''''''''''
Public Property Get soort() As String
    soort = pSoort
End Property
Public Property Let soort(Value As String)
    pSoort = Value
End Property

''''''''''''''''''''''
' StartDatum property
''''''''''''''''''''''
Public Property Get startdatum() As Date
    startdatum = pStartDatum
End Property
Public Property Let startdatum(Value As Date)
    pStartDatum = Value
End Property

''''''''''''''''''''''
' EindDatum property
''''''''''''''''''''''
Public Property Get einddatum() As Date
    einddatum = pEindDatum
End Property
Public Property Let einddatum(Value As Date)
    pEindDatum = Value
End Property
