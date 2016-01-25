VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blad3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public oudewaardeblad As Variant
Public startdatum As Date
Public einddatum As Date
Public update As Boolean
Public ishidden As Boolean

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim Status As Boolean
Dim logtekst As String
Dim t As taak
Dim synergy As String
Dim nieuwe_omschrijving As String

'Dim l As New Log
Blad3.oudewaardeblad = Target.Value
Dim ws As String
ws = Blad3.Name
If ThisWorkbook.Sheets(Blad3.Name).Range(SoortPlanning.col_taak_status & Target.Row) = "J" Then Status = True

'dubbelklik voor startdatum
If Not Intersect(Target, Range(SoortPlanning.col_taak_startdatum & ":" & SoortPlanning.col_taak_startdatum)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    If Range(SoortPlanning.col_taak_omschrijving & ActiveCell.Row) <> "" Then
        Set t = New taak
        t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row))
    End If
    
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_startdatum & Target.Row)
    ThisWorkbook.infokalender = "wijzig startdatum van taak: " & t.Omschrijving
    FORM_KALENDER.Show
    Application.EnableEvents = False
    
    If ThisWorkbook.inladen = True Then
        Target.NumberFormat = "m/d/yyyy"
        Target = ThisWorkbook.datum
        t.startdatum = ThisWorkbook.datum
        SoortPlanning.taakbalkverwijderen (Target.Row)
        If t.update Then
            Call SoortPlanning.TaakBalkPlaatsen(Lijsten.KalenderOverallPlanning, t.startdatum, t.einddatum, Status, ActiveCell.Row, True, ws)
            'logtekst = "Synergy: " & ThisWorkbook.Sheets(ws).Range("A" & Target.Row) & " / " & t.soort & " / "
            'logtekst = logtekst & t.omschrijving & " / Startdatum aangepast van " & CStr(oudewaardeblad1) & " in " & Target.value
            'l.createLog logtekst, tk_updaten, t.id, taak
        End If
    End If
End If

'dubbelklik voor TAAK.eindatum
If Not Intersect(Target, Range(SoortPlanning.col_taak_einddatum & ":" & SoortPlanning.col_taak_einddatum)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    If Range(SoortPlanning.col_taak_omschrijving & ActiveCell.Row) <> "" Then
        Set t = New taak
        t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row))
    End If
    
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_einddatum & Target.Row)
    ThisWorkbook.infokalender = "wijzig startdatum van taak: " & t.Omschrijving
    FORM_KALENDER.Show
    Application.EnableEvents = False
    
    If ThisWorkbook.inladen = True Then
        Target.NumberFormat = "m/d/yyyy"
        Target = ThisWorkbook.datum
        t.einddatum = ThisWorkbook.datum
        SoortPlanning.taakbalkverwijderen (Target.Row)
        If t.update Then
            Call SoortPlanning.TaakBalkPlaatsen(Lijsten.KalenderOverallPlanning, t.startdatum, t.einddatum, Status, ActiveCell.Row, True, ws)
            'logtekst = "Synergy: " & ThisWorkbook.Sheets(ws).Range("A" & Target.Row) & " / " & t.soort & " / "
            'logtekst = logtekst & t.omschrijving & " / Startdatum aangepast van " & CStr(oudewaardeblad1) & " in " & Target.value
            'l.createLog logtekst, tk_updaten, t.id, taak
        End If
    End If
End If

'dubbelklik omschrijving
If Not Intersect(Target, Range(SoortPlanning.col_taak_omschrijving & ":" & SoortPlanning.col_taak_omschrijving)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    Application.EnableEvents = False
    synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)
    Set t = New taak
    t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row))
    nieuwe_omschrijving = InputBox(synergy & vbNewLine & t.Omschrijving, "OMSCHRIJVING AANPASSEN", t.Omschrijving)
    If nieuwe_omschrijving <> "" Then
        t.Omschrijving = nieuwe_omschrijving
        If t.update = True Then
            ActiveCell.Value = nieuwe_omschrijving
        End If
    End If
    'logtekst = "Synergy: " & ThisWorkbook.Sheets(ws).Range("A" & Target.Row) & " / " & t.soort & " / "
    'logtekst = logtekst & t.omschrijving & " / Omschrijving aangepast van " & CStr(oudewaardeblad1) & " in " & Target.value
    'l.createLog logtekst, tk_updaten, t.id, taak
End If

Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim synergy As String
Dim opdracht As Variant
Dim Id As Long
Dim datum As Date
Dim duur As Double
Dim Status As Boolean
Dim Opmerking As String
Dim t As taak
'Dim l As New Log
Dim logtekst As String
Dim ws As String
Dim db As New DataBase
Dim Vestiging As String
Dim Waarde As String
Dim a As Long

ws = Blad3.Name

'aanpassen aantal
If Not Intersect(Target, Range(SoortPlanning.col_taak_duur & ":" & SoortPlanning.col_taak_duur)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    Set t = New taak
    t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & Target.Row))
    t.Aantal = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_duur & Target.Row)
    If t.update = True Then
'    logtekst = "Synergy: " & ThisWorkbook.Sheets(Blad3.Name).Range("A" & Target.Row) & " / " & t.soort & " / "
'    logtekst = logtekst & t.omschrijving & " / Aantal aangepast van " & CStr(oudewaardeblad8) & " in " & Target.value
'    l.createLog logtekst, tk_updaten, t.id, taak
    End If
End If

'aanpassen status
If Not Intersect(Target, Range(SoortPlanning.col_taak_status & ":" & SoortPlanning.col_taak_status)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    Set t = New taak
    t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & Target.Row))
    If ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_status & Target.Row) = "J" Then Status = True
    t.Status = Status
    If t.update = True Then
        SoortPlanning.taakbalkverwijderen (Target.Row)
        SoortPlanning.TaakBalkPlaatsen Lijsten.KalenderOverallPlanning, t.startdatum, t.einddatum, t.Status, Target.Row, True, ws
        If Status = True Then Target.Interior.Color = kleuren.taak_gereed Else Target.Interior.Color = kleuren.taak_niet_gereed
'    logtekst = "Synergy: " & ThisWorkbook.Sheets(Blad3.Name).Range("A" & Target.Row) & " / " & t.soort & " / "
'    logtekst = logtekst & t.omschrijving & " / Aantal aangepast van " & CStr(oudewaardeblad8) & " in " & Target.value
'    l.createLog logtekst, tk_updaten, t.id, taak
    End If
End If

'aanpassen opmerking
If Not Intersect(Target, Range(SoortPlanning.col_taak_opmerking & ":" & SoortPlanning.col_taak_opmerking)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    Set t = New taak
    t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & Target.Row))
    t.Opmerking = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_opmerking & Target.Row)
    t.Status = Status
    If t.update = True Then
'    logtekst = "Synergy: " & ThisWorkbook.Sheets(Blad3.Name).Range("A" & Target.Row) & " / " & t.soort & " / "
'    logtekst = logtekst & t.omschrijving & " / Aantal aangepast van " & CStr(oudewaardeblad8) & " in " & Target.value
'    l.createLog logtekst, tk_updaten, t.id, taak
    End If
End If


If Not Intersect(Target, Range(SoortPlanning.col_proj_intern & ":" & SoortPlanning.col_proj_intern)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(Blad3.Name).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)
    Vestiging = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & Target.Row)
    Waarde = ThisWorkbook.Sheets(Blad3.Name).Range(SoortPlanning.col_proj_intern & Target.Row)
    a = db.UpdateQueryUitvoeren("UPDATE PROJECTEN SET INTERN = '" & Waarde & "' WHERE SYNERGY = '" & synergy & "' AND Vestiging = '" & Vestiging & "';")
    If a = 0 Then MsgBox "Bijwerken van status Ingepland Intern is niet bijgewerkt.", vbCritical, "FOUUT AANPASSEN INGEPLAND INTERN"
End If

If Not Intersect(Target, Range(SoortPlanning.col_proj_extern & ":" & SoortPlanning.col_proj_extern)) Is Nothing And IsNumeric(ThisWorkbook.Sheets(Blad3.Name).Range(SoortPlanning.col_proj_synergy & Target.Row)) Then
    synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Target.Row)
    Vestiging = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & Target.Row)
    Waarde = ThisWorkbook.Sheets(Blad3.Name).Range(SoortPlanning.col_proj_extern & Target.Row)
    a = db.UpdateQueryUitvoeren("UPDATE PROJECTEN SET EXTERN = '" & Waarde & "' WHERE SYNERGY = '" & synergy & "' AND Vestiging = '" & Vestiging & "';")
    If a = 0 Then MsgBox "Bijwerken van status Ingepland Intern is niet bijgewerkt.", vbCritical, "FOUUT AANPASSEN INGEPLAND INTERN"
End If

End Sub

Public Sub TaakViewProd()
Dim ws As String
ws = Blad3.Name

If Blad3.ishidden = True Then
    Range(SoortPlanning.col_proj_synergy & ":" & SoortPlanning.col_plan_eindtijd).EntireColumn.Hidden = False
    Blad3.ishidden = False
    ThisWorkbook.Sheets(ws).Outline.ShowLevels RowLevels:=1
    ThisWorkbook.Sheets(ws).Outline.ShowLevels Columnlevels:=1
    
Else
    ThisWorkbook.Sheets(ws).Outline.ShowLevels Columnlevels:=3
    ThisWorkbook.Sheets(ws).Outline.ShowLevels RowLevels:=2
    Range(SoortPlanning.col_proj_opdrachtgever & ":" & SoortPlanning.col_plan_eindtijd).EntireColumn.Hidden = True
    Range(SoortPlanning.col_proj_PL & ":" & SoortPlanning.col_proj_PL).EntireColumn.Hidden = False
    Range(SoortPlanning.col_proj_WVB & ":" & SoortPlanning.col_proj_WVB).EntireColumn.Hidden = False
    Blad3.ishidden = True
End If
End Sub
