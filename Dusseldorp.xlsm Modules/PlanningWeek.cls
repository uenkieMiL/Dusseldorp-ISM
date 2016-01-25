VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanningWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pSynergy As String
Private pPV As String
Private pPL As String
Private pCAL As String
Private pWVB As String
Private pUITV As String
Private pOmschrijving As String
Private pOpdrachtgever As String
Private pKolomStart As Long
Private pKolomEind As Long
Private pKleur As Long
Private pSoort As String
Private pStartDatum As Date
Private pEindDatum As Date
Private pVestiging As String
Private pCProductie As Collection


''''''''''''''''''''''
' Synergy property
''''''''''''''''''''''
Public Property Get synergy() As String
    synergy = pSynergy
End Property
Public Property Let synergy(Value As String)
    pSynergy = Value
End Property

''''''''''''''''''''''
' PV property
''''''''''''''''''''''
Public Property Get pv() As String
    pv = pPV
End Property
Public Property Let pv(Value As String)
    pPV = Value
End Property

''''''''''''''''''''''
' PL property
''''''''''''''''''''''
Public Property Get pl() As String
    pl = pPL
End Property
Public Property Let pl(Value As String)
    pPL = Value
End Property

''''''''''''''''''''''
' CAL property
''''''''''''''''''''''
Public Property Get CAL() As String
    CAL = pCAL
End Property
Public Property Let CAL(Value As String)
    pCAL = Value
End Property

''''''''''''''''''''''
' WVB property
''''''''''''''''''''''
Public Property Get wvb() As String
    wvb = pWVB
End Property
Public Property Let wvb(Value As String)
    pWVB = Value
End Property

''''''''''''''''''''''
' UITV property
''''''''''''''''''''''
Public Property Get uitv() As String
    uitv = pUITV
End Property
Public Property Let uitv(Value As String)
    pUITV = Value
End Property

''''''''''''''''''''''
' Omschrijving property
''''''''''''''''''''''
Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property
Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property

''''''''''''''''''''''
' Opdrachtgever property
''''''''''''''''''''''
Public Property Get Opdrachtgever() As String
    Opdrachtgever = pOpdrachtgever
End Property
Public Property Let Opdrachtgever(Value As String)
    pOpdrachtgever = Value
End Property

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

''''''''''''''''''''''
' Vestiging property
''''''''''''''''''''''
Public Property Get Vestiging() As String
    Vestiging = pVestiging
End Property
Public Property Let Vestiging(Value As String)
    pVestiging = Value
End Property


'''''''''''''''''''''''''
' CBehoeften property
'''''''''''''''''''''''''

Public Property Get CProductie() As Collection
  Set CProductie = pCProductie
End Property

Public Property Let CProductie(Value As Collection)
  Set pCProductie = Value
End Property


Private Sub class_initialize()
    Set pCProductie = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCProductie = Nothing
End Sub

Property Get Count() As Long
    Count = pCProductie.Count
End Property

Property Get item(ByVal vID As Variant) As project
    Set item = pCProductie(vID)
End Property
Public Sub ToevoegenProductie(pr As PlanningWeekProductie)
    Dim p As New PlanningWeekProductie
    Set p = pr
    pCProductie.Add p
End Sub

