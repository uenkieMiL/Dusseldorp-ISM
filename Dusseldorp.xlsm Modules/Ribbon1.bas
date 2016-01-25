Attribute VB_Name = "Ribbon1"
Option Explicit

Dim Rib As IRibbonUI
Public MyTag As String
'Callback for customButton1.1 onAction
Sub Vernieuwenknop(control As IRibbonControl)
Dim ws As String
ws = ActiveSheet.Name
If ws = Blad1.Name Then
    SoortPlanning.MaakSoortPlanning ws, 1
ElseIf ws = Blad2.Name Then
    SoortPlanning.MaakSoortPlanning ws, 2
ElseIf ws = Blad3.Name Then
    SoortPlanning.MaakSoortPlanning ws, 4
ElseIf ws = Blad5.Name Then
    PersoneelsPlanning.MaakPersoneelsPlanning
ElseIf ws = Blad6.Name Then
    PPP.MaakProjectPersoneelsPlanning
End If


End Sub

'Callback for customButton1.3 onAction
Sub PlanningenToevoegen(control As IRibbonControl)
    FORM_PROJECT_AANMAKEN.Show
End Sub

'Callback for customButton1.4 onAction
Sub PlanningenAanpassen(control As IRibbonControl)
Dim Id As String
Dim Vestiging As String
Dim tekst As String
Dim V As Vestiging
Dim lijst As Collection
Set lijst = Lijsten.MaakLijstVestigingen

If ActiveSheet.Name = Blad1.Name Or ActiveSheet.Name = Blad2.Name Or ActiveSheet.Name = Blad3.Name Then
    ThisWorkbook.synergy_id = ThisWorkbook.Sheets(ActiveSheet.Name).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
    ThisWorkbook.Vestiging = ThisWorkbook.Sheets(ActiveSheet.Name).Range(SoortPlanning.col_proj_vestiging & ActiveCell.Row)
End If

Id = ThisWorkbook.synergy_id
Vestiging = ThisWorkbook.Vestiging

If Id = "" Then
    Id = InputBox("Geef AUB het synergy nummmer op!", "SYNERGY NUMMER OPGEVEN")
    tekst = "Geef AUB de vestiging op waar het om gaat door het nummerieke waarde in te geven:" & vbNewLine
    For Each V In lijst
        tekst = tekst & V.Id & " = " & V.Omschrijving & vbNewLine
    Next V
    Vestiging = InputBox(tekst, "SYNERGY NUMMER OPGEVEN")
End If

If Id = "" Then Exit Sub
If Vestiging = "" Then Exit Sub

If IsNumeric(Vestiging) = True Then
    Set V = lijst.item(Vestiging)
    Vestiging = V.Omschrijving
End If

If Functies.CheckProjectIsAangemaakt(Id, Vestiging) = False Then
    MsgBox "Synergy nummer is onbekend.", vbCritical, "SYNERGY NUMMER ONBEKEND"
    Exit Sub
End If

ThisWorkbook.synergy_id = Id
ThisWorkbook.Vestiging = Vestiging
FORM_PROJECT_WIJZIGEN.Show
End Sub

'Callback for customButton1.5 onAction
Sub WeekOverzichtMaken(control As IRibbonControl)
FORM_WEEKOVERZICHT.Show
End Sub

'Callback for customButton1.6 onAction
Sub GaNaarVandaag(control As IRibbonControl)
Dim ws As String
ws = ActiveSheet.Name

If ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name Then
    SoortPlanning.SelecteerKolomVandaag Lijsten.KalenderOverallPlanning
End If

End Sub

'Callback for customButton1.10 onAction
Sub ProjectOverzichtRibbon(control As IRibbonControl)
FORM_ALLE_PROJECTEN.Show

End Sub

'Callback for customButton1.7 onAction
Sub ZetInWacht(control As IRibbonControl)
Dim synergy As String
Dim Vestiging As String
Dim datumtekst As String
Dim datum As Date
Dim aantalwerkdagen As Long
Dim p As project
Dim ws As String

ws = ActiveSheet.Name
If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub

synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
If synergy <> "" And IsNumeric(synergy) = True Then
    Vestiging = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & ActiveCell.Row)
    Set p = New project
    p.synergy = synergy
    p.Vestiging = Vestiging
    p.haalop
    
    ThisWorkbook.inladen = False
    ThisWorkbook.infokalender = "Zet project " & p.synergy & vbNewLine & "Vestging " & p.Vestiging & " in de wacht"
    FORM_KALENDER.Show
    If IsDate(ThisWorkbook.datum) = True And ThisWorkbook.inladen = True Then
        
        p.naBelDatum = ThisWorkbook.datum
        p.staatInWacht = True
        If p.UpdateWacht = True Then SoortPlanning.VerwijderProject
        'l.createLog "Project in de wacht gezet, Nabeldatum = " & CStr(ThisWorkbook.datum), pr_updaten, synergy, project
        
    End If

End If
End Sub

'Callback for customButton1.8 onAction
Sub HaalUitWacht(control As IRibbonControl)
FORM_STAATINWACHT.Show
End Sub

'Callback for customButton1.9 onAction
Sub BeheerKalender(control As IRibbonControl)
FORM_KALENDER_BEHEREN.Show
End Sub

'Callback for customButton1.11 onAction
Sub VerwijderProjectRibbon(control As IRibbonControl)
Dim ws As String
Dim p As New project
Dim synergy As String
Dim Vestiging As String

ws = ActiveSheet.Name

If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub
synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
Vestiging = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & ActiveCell.Row)
If IsNumeric(synergy) = False Or Vestiging = "" Then Exit Sub

p.synergy = synergy
p.Vestiging = Vestiging
p.haalop
p.verwijderenProject
SoortPlanning.VerwijderProject
End Sub

'Callback for customButton2.1 onAction
Sub ActieToevoegen(control As IRibbonControl)
Dim synergy As String
Dim veld As String
Dim opdracht As Variant
Dim opdrachtmax As Variant
Dim offsetrij As Long
Dim rij As Long
Dim t As New taak
Dim ws As String
Dim CKalender As New Collection
Dim db As New DataBase

ws = ActiveSheet.Name
If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub
If IsNumeric(ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)) And _
Not IsEmpty(ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_taak_soort & ActiveCell.Row)) Then

        t.haalop (ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row))
        t.CopyTaak
        Turbo_AAN
        offsetrij = t.Volgnummer - ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_taak_volgnummer & ActiveCell.Row)
        rij = ActiveCell.Row
        ActiveCell.Offset(1).EntireRow.insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_vestiging & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_omschrijving & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_omschrijving & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_opdrachtgever & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_opdrachtgever & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_PV & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_PV & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_PL & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_PL & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_CALC & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_CALC & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_WVB & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_WVB & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_UITV & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_UITV & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_plan_starttijd & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_plan_starttijd & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_plan_eindtijd & rij + 1) = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_plan_eindtijd & ActiveCell.Row)
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_omschrijving & rij + 1) = t.Omschrijving
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_volgnummer & rij + 1) = t.Volgnummer
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_startdatum & rij + 1) = t.startdatum
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_einddatum & rij + 1) = t.einddatum
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_duur & rij + 1) = t.Aantal
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_ehd & rij + 1) = t.Ehd
        If t.Status = True Then ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_status & rij + 1) = "J" Else ThisWorkbook.Sheets(ws).Range("T" & rij + 1) = "N"
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & rij + 1) = t.Id
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_veld & rij + 1) = t.veld
        ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_soort & rij + 1) = t.soort
        Set CKalender = Lijsten.KalenderOverallPlanning
        Functies.TaakBalkPlaatsen CKalender, t.startdatum, t.einddatum, t.Status, rij + 1, True
        turbo_UIT
End If
End Sub

'Callback for customButton2.2 onAction
Sub ActieVerwijderen(control As IRibbonControl)
Dim t As taak
Dim Id As Long
Dim synergy As String, veld As String
Dim ws As String

ws = ActiveSheet.Name
If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub

    If IsNumeric(ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)) And _
    Not IsEmpty(ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_taak_soort & ActiveCell.Row)) Then
        synergy = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
        Id = ThisWorkbook.Worksheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row)
            Set t = New taak
            t.haalop (Id)
            If t.delete = True Then
                Turbo_AAN
                Rows(ActiveCell.Row & ":" & ActiveCell.Row).delete Shift:=xlUp
                turbo_UIT
            End If
    End If

End Sub

'Callback for customButton2.3 onAction
Sub VrijeActieToevoegen(control As IRibbonControl)
Dim p As New project
Dim t As New taak
Dim taak As taak
Dim synergy As String
Dim Omschrijving As String
Dim pl As Planning
Dim datum As Date
Set taak = New taak
Dim rij As Long
Dim ws As String
Dim CKalender As New Collection
Dim planid As Long
Dim Vestiging As String
Dim soortnaam As String

ws = ActiveSheet.Name
If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub

synergy = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
Vestiging = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_vestiging & ActiveCell.Row)
Omschrijving = InputBox("Geef een omschrijving van de taak.", "VRIJE TAAK TOEVOEGEN")
If Omschrijving = "" Then Exit Sub
p.synergy = synergy
p.Vestiging = Vestiging
p.haalop
If ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_omschrijving & ActiveCell.Row) = "" Then
    planid = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row)
Else
    t.haalop ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & ActiveCell.Row)
    planid = t.planningid
End If

For Each pl In p.CPlanningen
    soortnaam = pl.SoortByteNaarKortStringTerug
    If soortnaam = t.soort Then
        datum = pl.startdatum
        Exit For
    End If
    
Next pl


taak.planningid = planid
taak.Omschrijving = Omschrijving
taak.Volgnummer = 1
taak.startdatum = datum
taak.einddatum = datum
taak.Aantal = 1
taak.Ehd = "uur"
taak.Status = False
taak.veld = "ZZVRIJ"
taak.soort = t.soort
taak.BegrotingsRegel = False
taak.VoegTaakToe

If ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row + 1) = "" Then
    rij = ActiveCell.Row + 1
Else
    rij = ActiveCell.End(xlDown).Row + 1
End If


Application.EnableEvents = False

Rows(rij & ":" & rij).insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & rij) = CLng(synergy)
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_omschrijving & rij) = Omschrijving
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_volgnummer & rij) = taak.Volgnummer
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_startdatum & rij) = taak.startdatum
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_einddatum & rij) = taak.einddatum
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_duur & rij) = taak.Aantal
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_ehd & rij) = taak.Ehd
If taak.Status = True Then ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_status & rij) = "J" Else ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_status & rij) = "N"
Set CKalender = Lijsten.KalenderOverallPlanning
Functies.TaakBalkPlaatsen CKalender, t.startdatum, t.einddatum, t.Status, rij, True
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & rij) = taak.Id
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_veld & rij) = taak.veld
ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_soort & rij) = taak.soort
Application.EnableEvents = True
End Sub

'Callback for customButton2.4 onAction
Sub TaakbalkAanpassenSelectie(control As IRibbonControl)
Dim t As taak
Dim ws As String

ws = ActiveSheet.Name
If Not (ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name) Then Exit Sub

    If IsNumeric(ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & Selection.Row)) = True Then
       If ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_omschrijving & Selection.Row) <> "" Then
           Set t = New taak
           t.haalop (ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & Selection.Row))
            SelecteerDataVanSelectie Range(Selection.Address), t
       End If
    End If


End Sub

'Callback for customButton2.5 onAction
Sub VerzettenOpdrachten(control As IRibbonControl)
Dim ws As String

ws = ActiveSheet.Name

If ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name Then
    SoortPlanning.TakenBinnenPlanningAantalDagenDoorschuiven
End If

End Sub

'Callback for customButton2.6 onAction
Sub RegieTaken(control As IRibbonControl)
End Sub

'Callback for customButton2.7 onAction
Sub PersoneelBeheren(control As IRibbonControl)
FORM_PERSONEEL_OVERZICHT.Show
End Sub

'Callback for customButton3.3 onAction
Sub ProjectMapOpenen(control As IRibbonControl)
Dim synergy As String
'Dim l As New Log
Dim ws As String

ws = ActiveSheet.Name
If ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name Then

    synergy = ThisWorkbook.Sheets(ActiveSheet.Name).Range("A" & ActiveCell.Row)
    If synergy <> "" And IsNumeric(synergy) = True Then
    Functies.OpenFolder (synergy)
        'l.createLog "Project map openen " & synergy, overzicht_gemaakt, "PROJECTMAP OPENEN", 5
    End If
End If
End Sub
'Callback for customButton4.01 onAction
Sub AanmakenAgendaItem(control As IRibbonControl)
Dim ws As String

ws = ActiveSheet.Name
If ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name Then
    If IsNumeric(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)) = True Then
        Call AgendaItemToevoegen
    End If
End If

End Sub


'Callback for customButton4.02 onAction
Sub AanmakenTaakItem(control As IRibbonControl)
Dim ws As String

ws = ActiveSheet.Name
If ws = Blad1.Name Or ws = Blad2.Name Or ws = Blad3.Name Then
    If IsNumeric(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)) = True Then
        Call TaakItemToevoegen
    End If
End If

End Sub

'Callback for customButton4.03 onAction
Sub MaakMail(control As IRibbonControl)
    Outlook.Mail_ActiveSheet
End Sub




'Callback for customUI.onLoad
Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set Rib = ribbon
    'If you want to run a macro below when you open the workbook
    'you can call the macro like this :
    'Call EnableControlsWithCertainTag3
End Sub

Sub GetEnabledMacro(control As IRibbonControl, ByRef Enabled)
    If MyTag = "Enable" Then
        Enabled = True
    Else
        If control.Tag Like MyTag Then
            Enabled = True
        Else
            Enabled = False
        End If
    End If
End Sub

Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
        MsgBox "Error, Save/Restart your workbook" & vbNewLine & _
        "Visit this page for a solution: http://www.rondebruin.nl/ribbonstate.htm"
    Else
        Rib.Invalidate
    End If
End Sub

'Callback for customButton4.1 onAction
Sub MaterieelBestellen(control As IRibbonControl)
Dim ws As New Worksheet

Set ws = ActiveSheet
If ws.Name = Blad1.Name Or ws.Name = Blad2.Name Or ws.Name = Blad3.Name Then
    If IsNumeric(ws.Range("A" & ActiveCell.Row)) = True And ws.Range("A" & ActiveCell.Row) <> "" Then
        ThisWorkbook.synergy_id = ws.Range("A" & ActiveCell.Row)
        ThisWorkbook.Vestiging = ws.Range("B" & ActiveCell.Row)
        FORM_MATERIEEL_BESTELLEN.Show
    End If
End If
    
End Sub

