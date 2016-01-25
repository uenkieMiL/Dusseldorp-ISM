VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanningSoort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private Pnaam As String
Dim pCProjecten As Collection

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
' CProjecten property
'''''''''''''''''''''''''

Public Property Get CProjecten() As Collection
    Set CProjecten = pCProjecten
End Property

Public Property Let CProjecten(Value As Collection)
   Set pCProjecten = Value
End Property


Property Get Count() As Long
    Count = pCProjecten.Count
End Property

Property Get item(ByVal vID As Variant) As project
    Set item = pCProjecten(vID)
End Property

Private Sub class_initialize()
    Set pCProjecten = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCProjecten = Nothing
End Sub

Public Sub ToevoegenProject(ByRef p As project)
    Dim ap As project
    Set ap = New project
    Set ap = p
    pCProjecten.Add p
End Sub
