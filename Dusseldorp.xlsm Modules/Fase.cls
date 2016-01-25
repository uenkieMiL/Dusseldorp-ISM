VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Byte
Private pOmschrijving As String
Private pCProjecten As Collection


'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get Id() As Byte
    Id = pId
End Property

Public Property Let Id(Value As Byte)
    pId = Value
End Property

'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property

'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get CProjecten() As Collection
    Set CProjecten = New Collection
    Set CProjecten = pCProjecten
End Property

Public Property Let CProjecten(Value As Collection)
   Set pCProjecten = Value
End Property


Private Sub class_initialize()
    Set pCProjecten = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCProjecten = Nothing
End Sub
Property Get Count() As Long
    Count = pCProjecten.Count
End Property

Property Get item(ByVal vID As Variant) As PlanningWeek
    Set item = pCProjecten(vID)
End Property
Public Sub ToevogenPlanningsweek(ByRef pw As PlanningWeek)
    Dim p As New PlanningWeek
    Set p = pw
    pCProjecten.Add p
End Sub

