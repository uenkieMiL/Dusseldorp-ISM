VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public inladen As Boolean
Public datum As Date
Public infokalender As String
Public synergy_id As String
Public Vestiging As String
Public mat_id As Long

Private Sub Workbook_Open()
    'Dim Log As New Log
    Dim antwoord As Integer
    
    If Environ("UserName") = "Roderik" Then
        antwoord = MsgBox("Programma Normaal Opstarten?", vbYesNo, "OPSTARTEN")
        If antwoord = vbNo Then Exit Sub
    End If
    
    'Log.createLog "Openen Bestand", app_openen, "NVT", RelatieEnum.Overig
    Turbo_AAN
        SoortPlanning.MaakSoortPlanning Blad1.Name, 1
        SoortPlanning.MaakSoortPlanning Blad2.Name, 2
        SoortPlanning.MaakSoortPlanning Blad3.Name, 4
        PersoneelsPlanning.MaakPersoneelsPlanning
        MaakProjectPersoneelsPlanning
    turbo_UIT
End Sub
