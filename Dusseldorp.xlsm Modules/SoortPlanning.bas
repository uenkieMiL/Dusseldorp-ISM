Attribute VB_Name = "SoortPlanning"
Option Explicit
Private CKalender As Collection
Public Const startkolom As Integer = 25
Public Const startrij As Integer = 6
Private Const group1Rng As String = "L:U"
Private Const group2Rng As String = "G:U"
Private Const hiddenRng As String = "V:X"
Public Const col_proj_synergy As String = "A"
Public Const col_proj_vestiging As String = "B"
Public Const col_proj_omschrijving As String = "C"
Public Const col_proj_opdrachtgever As String = "D"
Public Const col_proj_intern As String = "E"
Public Const col_proj_extern As String = "F"
Public Const col_proj_PV As String = "G"
Public Const col_proj_PL As String = "H"
Public Const col_proj_CALC As String = "I"
Public Const col_proj_WVB As String = "J"
Public Const col_proj_UITV As String = "K"
Public Const col_plan_starttijd As String = "L"
Public Const col_plan_eindtijd As String = "M"
Public Const col_taak_omschrijving As String = "N"
Public Const col_taak_volgnummer As String = "O"
Public Const col_taak_startdatum As String = "P"
Public Const col_taak_einddatum As String = "Q"
Public Const col_taak_duur As String = "R"
Public Const col_taak_ehd As String = "S"
Public Const col_taak_status As String = "T"
Public Const col_taak_opmerking As String = "U"
Public Const col_id As String = "V"
Public Const col_veld As String = "W"
Public Const col_taak_soort As String = "X"
Public Const col_begin_plan As String = "Y"

Public Const row_datum As Long = 1
Public Const row_jaar As Long = 2
Public Const row_maand As Long = 3
Public Const row_week As Long = 4
Public Const row_dag As Long = 5



Function MaakSoortPlanning(werkblad As String, PlanningSoort As Byte)

Dim d As datum
Dim laatsterij As Integer 'laatste rij
Dim laatstekolom As Integer 'laatste kolom
Dim rng As Range

Turbo_AAN


If werkblad <> ActiveSheet.Name Then Sheets(werkblad).Select
Range("A1").Select
If Sheets(werkblad).AutoFilterMode = True Then Sheets(werkblad).AutoFilterMode = False

If CKalender Is Nothing Then Set CKalender = Lijsten.KalenderOverallPlanning

ThisWorkbook.Sheets(werkblad).Outline.ShowLevels RowLevels:=3
ThisWorkbook.Sheets(werkblad).Cells.Rows.ClearOutline
ThisWorkbook.Sheets(werkblad).Columns(group1Rng).Group
ThisWorkbook.Sheets(werkblad).Columns(group2Rng).Group

ThisWorkbook.Sheets(werkblad).Columns(hiddenRng).EntireColumn.Hidden = True

Sheets(werkblad).Outline.ShowLevels RowLevels:=0, Columnlevels:=2
laatsterij = ThisWorkbook.Sheets(werkblad).Range("N100000").End(xlUp).Row
laatstekolom = ThisWorkbook.Sheets(werkblad).Range("XFD5").End(xlToLeft).Column
If laatstekolom >= startkolom Then ThisWorkbook.Sheets(werkblad).Range(Range(Columns(startkolom), Columns(laatstekolom + 1)).Address).delete Shift:=xlToLeft
If laatsterij > 4 Then ThisWorkbook.Sheets(werkblad).Range(Range(Cells(startrij, 1), Cells(laatsterij + 1, 16384)).Address).Clear

For Each d In CKalender
    If d.Kolomnummer > -1 Then
    ThisWorkbook.Sheets(werkblad).Range(Cells(1, startkolom + d.Kolomnummer).Address) = d.datum
    ThisWorkbook.Sheets(werkblad).Range(Cells(2, startkolom + d.Kolomnummer).Address) = Year((d.datum))
    ThisWorkbook.Sheets(werkblad).Range(Cells(3, startkolom + d.Kolomnummer).Address) = MonthName(Month(d.datum))
    ThisWorkbook.Sheets(werkblad).Range(Cells(4, startkolom + d.Kolomnummer).Address) = DatePart("ww", d.datum, vbMonday, vbFirstFourDays)
    ThisWorkbook.Sheets(werkblad).Range(Cells(5, startkolom + d.Kolomnummer).Address) = Day(d.datum)
    End If
Next d

Functies.DagWeekSamenvoegen werkblad, 4
Functies.DagMaandSamenvoegen werkblad, 3
Functies.DagJaarSamenvoegen werkblad, 2


laatstekolom = ThisWorkbook.Sheets(werkblad).Range("XFD5").End(xlToLeft).Column
' plaats dikkelijn onderaan rij 5
    With ThisWorkbook.Sheets(werkblad).Range(Range(Cells(startrij, 1), Cells(startrij, laatstekolom)).Address).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With

Call Functies.Kolombreedte(3, ThisWorkbook.Sheets(werkblad).Range(Range(Cells(4, startkolom), Cells(4, laatstekolom)).Address), werkblad)
Call Functies.MaakKolomVandaag(CKalender, werkblad, startkolom)
Call Functies.PlaatsFeestdagen(CKalender, werkblad, startkolom)
SoortPlanning.PlaatsDetails PlanningSoort, CKalender, werkblad
Functies.DikkeStrepen werkblad, CKalender, startkolom
ThisWorkbook.Sheets(werkblad).Outline.ShowLevels RowLevels:=1
ThisWorkbook.Sheets(werkblad).Outline.ShowLevels Columnlevels:=1
If werkblad = Blad3.Name Then setopmaakactiehouder
Set rng = ThisWorkbook.Sheets(werkblad).Range(Rows(1).Address)
rng.RowHeight = 0
turbo_UIT


End Function

Public Function terminate_Ckalender()
    Set SoortPlanning.CKalender = Nothing
End Function

Function setopmaakactiehouderregel(kolom As String, veld As String)
    With Range("$" & kolom & ":$" & kolom)
        .FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$" & col_veld & "1=" & """" & veld & """"
        With .FormatConditions(.FormatConditions.Count)
            .SetFirstPriority
            With .Interior
                .Color = kleuren.actiehouder
            End With
        End With
    End With
End Function

Function setopmaakactiehouder()
If ActiveSheet.Name <> Blad3.Name Then Blad3.Select
Blad3.Cells.FormatConditions.delete
Range("A1").Select
    
    setopmaakactiehouderregel col_proj_PL, "uitv01"
    setopmaakactiehouderregel col_proj_PL, "uitv02"
    setopmaakactiehouderregel col_proj_PL, "uitv03"
    setopmaakactiehouderregel col_proj_PL, "uitv05"
    setopmaakactiehouderregel col_proj_WVB, "uitv06"
    setopmaakactiehouderregel col_proj_WVB, "uitv07"
    setopmaakactiehouderregel col_proj_WVB, "uitv08"
    setopmaakactiehouderregel col_proj_PL, "uitv09"
    setopmaakactiehouderregel col_proj_WVB, "uitv10"
    setopmaakactiehouderregel col_proj_WVB, "uitv11"
    setopmaakactiehouderregel col_proj_WVB, "uitv12"
    setopmaakactiehouderregel col_proj_PL, "uitv13"
    setopmaakactiehouderregel col_proj_PL, "uitv14"
    setopmaakactiehouderregel col_proj_PL, "uitv15"
    
End Function

Function PlaatsDetails(soort As Byte, CKalender As Collection, ws As String)
Dim planner As Collection
Dim t As taak
Dim p As project
Dim r As Long: r = startrij
Dim rng As Range
Dim titel As Range
Dim synergy As String
Dim pr As Productie
Dim startr As Long
Dim d As datum
Dim k1 As Long
Dim k2 As Long
Dim lk As Long

Sheets(ws).Select

Set planner = Lijsten.MaakSoortPlanningv2(soort)
For Each p In planner

        If p.PlanningVanProject.soort = soort Then
            If ThisWorkbook.Sheets(ws).Range(col_proj_vestiging & r) <> p.Vestiging And r <> startrij Then
                r = r + 1
                ThisWorkbook.Sheets(ws).Range(r & ":" & r).Interior.Color = 1
                Rows(r & ":" & r).RowHeight = 15
            End If
           
            
            
            If r = 3 Then r = r + 2 Else r = r + 1
            
            synergy = p.synergy
            ThisWorkbook.Sheets(ws).Range(col_id & r) = p.PlanningVanProject.Id
            ThisWorkbook.Sheets(ws).Range(col_veld & r) = p.PlanningVanProject.soort
            
            ThisWorkbook.Sheets(ws).Range("A" & r & ":" & col_taak_opmerking & r).Interior.Color = 15921906
            On Error Resume Next
                Set d = New datum
                Set d = CKalender.item(CStr(p.PlanningVanProject.startdatum))
                If d.datum = "0:00:00" Then k1 = -1 Else k1 = d.Kolomnummer
                Set d = New datum
                Set d = CKalender.item(CStr(p.PlanningVanProject.einddatum))
                If d.datum = "0:00:00" Then k2 = -1 Else k2 = d.Kolomnummer
            Resume Next
            If k1 > -1 And k2 > -1 Then
                Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, startkolom + k1), Cells(r, startkolom + k2))
            ElseIf k1 = -1 And k2 = -1 Then
            Set rng = Nothing
            Else
                If k1 = -1 Then k1 = 0
                If k2 = -1 Then k2 = 0
                Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, startkolom + k1), Cells(r, startkolom + k2))
            End If
            If Not rng Is Nothing Then
                If soort <> 4 Then
                    rng.Interior.Color = 0
                End If
            End If
            If soort = 4 Then
                For Each pr In p.CProducties
                Call ProductieBalkPlaatsen(CKalender, pr.startdatum, pr.einddatum, pr.Kleur, r, ws)
                Next pr
            End If
            startr = r + 1
            
            For Each t In p.PlanningVanProject.cTaken
            r = r + 1
'
            With ThisWorkbook.Sheets(ws).Range(col_taak_omschrijving & r)
                If t.BegrotingsRegel = True Then
                    .InsertIndent 1
                    .Font.Italic = True
                    
                    If t.Artikelnummer <> "" Then
                        .InsertIndent 2
                        .Font.Italic = True
                    End If
                End If
                .Value = t.Omschrijving
            End With
            
            ThisWorkbook.Sheets(ws).Range(col_taak_volgnummer & r) = t.Volgnummer
            ThisWorkbook.Sheets(ws).Range(col_taak_startdatum & r) = t.startdatum
            ThisWorkbook.Sheets(ws).Range(col_taak_einddatum & r) = t.einddatum
            ThisWorkbook.Sheets(ws).Range(col_taak_duur & r) = t.Aantal
            ThisWorkbook.Sheets(ws).Range(col_taak_ehd & r) = t.Ehd
            MaakLijstJN (ThisWorkbook.Sheets(ws).Range(col_taak_status & r))
            If t.Status = True Then
                ThisWorkbook.Sheets(ws).Range(col_taak_status & r) = "J"
                ThisWorkbook.Sheets(ws).Range(col_taak_status & r).Interior.Color = 5287936
            Else
                ThisWorkbook.Sheets(ws).Range(col_taak_status & r) = "N"
                ThisWorkbook.Sheets(ws).Range(col_taak_status & r).Interior.Color = 192
            End If
            TaakBalkPlaatsen CKalender, t.startdatum, t.einddatum, t.Status, r, True, ws
            ThisWorkbook.Sheets(ws).Range(col_taak_opmerking & r) = t.Opmerking 'Opmerking
            ThisWorkbook.Sheets(ws).Range(col_id & r) = t.Id
            ThisWorkbook.Sheets(ws).Range(col_veld & r) = t.veld
            ThisWorkbook.Sheets(ws).Range(col_taak_soort & r) = t.soort
            ThisWorkbook.Sheets(ws).Range(col_proj_synergy & r) = synergy
            
            Next t
            
            
            Rows(startr & ":" & r).RowHeight = 15
            ThisWorkbook.Sheets(ws).Range(col_proj_synergy & startr - 1 & ":" & col_proj_synergy & r).Value = synergy
            ThisWorkbook.Sheets(ws).Range(col_proj_vestiging & startr - 1 & ":" & col_proj_vestiging & r) = p.Vestiging
            ThisWorkbook.Sheets(ws).Range(col_proj_omschrijving & startr - 1 & ":" & col_proj_omschrijving & r) = p.Omschrijving
            ThisWorkbook.Sheets(ws).Range(col_proj_opdrachtgever & startr - 1 & ":" & col_proj_opdrachtgever & r) = p.Opdrachtgever
            If soort = 4 Then
                MaakLijstIenE (ThisWorkbook.Sheets(ws).Range(col_proj_intern & startr - 1 & ":" & col_proj_intern & r))
                MaakLijstIenE (ThisWorkbook.Sheets(ws).Range(col_proj_extern & startr - 1 & ":" & col_proj_extern & r))
            End If
            ThisWorkbook.Sheets(ws).Range(col_proj_intern & startr - 1 & ":" & col_proj_intern & r) = p.intern
            ThisWorkbook.Sheets(ws).Range(col_proj_extern & startr - 1 & ":" & col_proj_extern & r) = p.extern
            ThisWorkbook.Sheets(ws).Range(col_proj_PV & startr - 1 & ":" & col_proj_PV & r) = p.pv
            ThisWorkbook.Sheets(ws).Range(col_proj_PL & startr - 1 & ":" & col_proj_PL & r) = p.pl
            ThisWorkbook.Sheets(ws).Range(col_proj_CALC & startr - 1 & ":" & col_proj_CALC & r) = p.CALC
            ThisWorkbook.Sheets(ws).Range(col_proj_WVB & startr - 1 & ":" & col_proj_WVB & r) = p.wvb
            ThisWorkbook.Sheets(ws).Range(col_proj_UITV & startr - 1 & ":" & col_proj_UITV & r) = p.uitv
            If p.CProducties.Count > 0 Then
            ThisWorkbook.Sheets(ws).Range(col_plan_starttijd & startr - 1 & ":" & col_plan_starttijd & r) = p.PlanningVanProject.startdatum
            ThisWorkbook.Sheets(ws).Range(col_plan_eindtijd & startr - 1 & ":" & col_plan_eindtijd & r) = p.PlanningVanProject.einddatum
            End If
            
            
           
            
            If p.PlanningVanProject.cTaken.Count > 0 Then
                ThisWorkbook.Sheets(ws).Range(startr & ":" & r).Group
            End If
        End If
        
        
        
    Next p
    
    ThisWorkbook.Sheets(ws).Range("A" & startrij & ":" & col_taak_opmerking & r).AutoFilter
    ThisWorkbook.Sheets(ws).Range(col_proj_intern & startrij & ":" & col_plan_eindtijd & r).HorizontalAlignment = xlCenter
    ThisWorkbook.Sheets(ws).Range(col_taak_omschrijving & startrij & ":" & col_taak_status & r).HorizontalAlignment = xlCenter
    
    
    lk = ThisWorkbook.Sheets(Blad1.Name).Range("XFD" & startrij - 1).End(xlToLeft).Column
    Functies.MaakRaster (ThisWorkbook.Sheets(ws).Range("A" & startrij & ":" & Cells(r, lk).Address))
    ThisWorkbook.Sheets(ws).Range(col_taak_omschrijving & startrij - 2) = SoortnaarString(soort)
    Set planner = Nothing
End Function

Function TaakBalkPlaatsen(CKalender As Collection, startdatum As Date, einddatum As Date, Status As Boolean, r As Long, uitvoeren As Boolean, ws As String)
        Dim d As datum
        Dim k1 As Long: k1 = -1
        Dim k2 As Long: k2 = -1
        Dim Kleur As Long
        Dim lk As Long
        Dim rng As Range
        ws = ActiveSheet.Name
        If Status = True Then Kleur = 5287936 Else Kleur = 192
        On Error Resume Next
            Set d = New datum
            Set d = CKalender.item(CStr(startdatum))
            If d.datum = "0:00:00" Then k1 = -1 Else k1 = d.Kolomnummer
            Set d = New datum
            Set d = CKalender.item(CStr(einddatum))
            If d.datum = "0:00:00" Then k2 = -1 Else k2 = d.Kolomnummer
        Resume Next
        If k2 = -1 And k1 = -1 Then Exit Function
        If einddatum = #12:00:00 AM# Then uitvoeren = False: k1 = 0: k2 = lk
        Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, startkolom + k1), Cells(r, startkolom + k2))
        If uitvoeren = True Then rng.Interior.Color = Kleur
        
End Function

Function taakbalkverwijderen(rij As Long)
    Dim ws As String
    Dim lk As Long
    Dim rng As Range
    
    ws = ActiveSheet.Name
    lk = SoortPlanning.laatstekolom
    Set rng = ThisWorkbook.Sheets(ws).Range(Cells(rij, startkolom), Cells(rij, lk))
    rng.Interior.Color = xlNone
End Function

Function ProductieBalkPlaatsen(CKalender As Collection, startdatum As Date, einddatum As Date, Kleur As Long, r As Long, ws As String)
        Dim d As datum
        Dim k1 As Long
        Dim k2 As Long
        Dim lk As Long
        Dim rng As Range
        Dim uitvoeren As Boolean
    
        For Each d In CKalender
            If startdatum = d.datum Then k1 = d.Kolomnummer
            If einddatum = d.datum Then k2 = d.Kolomnummer
            If k1 <> 0 And k2 <> 0 Then Exit For
            lk = d.Kolomnummer
        Next d
        If einddatum = #12:00:00 AM# Then k1 = 0: k2 = lk
        Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, startkolom + k1), Cells(r, startkolom + k2))
        rng.Interior.Color = Kleur
        
End Function


Function SelecteerKolomVandaag(CKalender As Collection)
Dim vandaag As Date: vandaag = Now(): vandaag = FormatDateTime(Date, vbShortDate)
Dim k As datum
Dim gevonden As Boolean
Dim kolom As Long
Dim rng As Range
Dim ws As String
ws = ActiveSheet.Name
Do While gevonden = False
    For Each k In CKalender
        If k.datum = vandaag Then
            kolom = k.Kolomnummer
            gevonden = True
            Exit For
        End If
    Next k
    
    If gevonden = False Then vandaag = DateAdd("d", 1, vandaag)
Loop

ThisWorkbook.Sheets(ws).Range(Cells(ActiveCell.Row, startkolom + kolom).Address).Select
End Function

Sub VerwijderProject()

Dim r As Long
    Dim a As Long
    Dim sr As Long
    Dim lr As Long
    Dim ws As String
    
    ws = ActiveSheet.Name
    r = ActiveCell.Row
    
    Do Until ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & r + a) <> ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
    a = a - 1
    Loop
    
    sr = r + a + 1
    
    
    a = 0
      Do Until ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & sr + a) <> ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row)
    a = a + 1
    Loop
    
    lr = sr + a - 1
    Turbo_AAN
    ThisWorkbook.Sheets(ws).Range(sr & ":" & lr).EntireRow.delete
    turbo_UIT
End Sub


Function SelecteerDataVanSelectie(r As Range, t As taak)
Dim CKalendar As New Collection
Dim eerstekolom As Long
Dim laatstekolom As Long
Dim startdatum As Date
Dim einddatum As Date
Dim a As Long
Dim Status As Boolean
'Dim l As New Log
Dim tekst As String
Dim oudestart As Date: oudestart = t.startdatum
Dim oudeeind As Date: oudeeind = t.einddatum
Dim oudebalk As Range
Dim ws As String
Turbo_AAN
ws = ActiveSheet.Name

Set CKalendar = Lijsten.KalenderOverallPlanning
verwijderoudebalk r.Row, t
eerstekolom = r.Column
a = r.Columns.Count
laatstekolom = eerstekolom + a - 1
t.startdatum = KolomNaarDatumviaNummer(eerstekolom, CKalendar)
If a = 1 Then t.einddatum = t.startdatum Else t.einddatum = KolomNaarDatumviaNummer(laatstekolom, CKalendar)

If t.update = True Then
    TaakBalkPlaatsenMulti eerstekolom, laatstekolom, t.Status, r.Row, True
    ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_startdatum & r.Row).Value = t.startdatum
    ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_einddatum & r.Row).Value = t.einddatum
        
'    tekst = "TAAKBALK AANGEPAST {"
'    tekst = tekst & "Startdatum van " & CStr(oudestart) & " in " & CStr(t.startdatum) & ", "
'    tekst = tekst & "Einddatum van " & CStr(oudeeind) & " in " & CStr(t.einddatum)
'    tekst = tekst & "}"
    'l.createLog tekst, tk_aanmaken, t.id, taak
End If

turbo_UIT
End Function

Function TakenBinnenPlanningAantalDagenDoorschuiven()
Dim Aantal As Integer
Dim planid As Long
Dim plansoort As Byte
Dim lijst As Variant
Dim t As New taak
Dim CKalender As Collection
Dim ws As String
Dim Waarde As String
Dim r As Long
Dim a As Long
ws = ActiveSheet.Name

Waarde = InputBox("Geeft het aantal dagen op dat je de planning wilt doorschuiven?", "DAGEN DOORSCHUIVEN", 0)
If Waarde = "" Then Exit Function
If IsNumeric(Waarde) = False Then
    MsgBox "De waarde die is ingevoerd is niet numeriek. de actie wordt geannuleerd.", vbCritical, "GEEN NUMERIEKE WAARDE"
    Exit Function
End If
Aantal = CLng(Waarde)

Turbo_AAN
Set CKalender = Lijsten.KalenderOverallPlanning
r = ActiveCell.Row

Do Until ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & r + a).Value <> ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row).Value
a = a - 1
Loop
a = a + 2
Do Until ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & r + a).Value <> ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_proj_synergy & ActiveCell.Row).Value
    
    If ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_status & r + a).Value <> "J" Then
        If ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_soort & r + a).Value <> "" Then
           
            planid = ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_id & r + a).Value
            t.haalop planid
            t.startdatum = DatumAantalWerkDagenVerplaatsenCollection(t.startdatum, Aantal, CKalender)
            t.einddatum = DatumAantalWerkDagenVerplaatsenCollection(t.einddatum, Aantal, CKalender)
            ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_startdatum & r + a).Value = t.startdatum
            ThisWorkbook.Sheets(ws).Range(SoortPlanning.col_taak_einddatum & r + a).Value = t.einddatum
            ThisWorkbook.Sheets(ws).Range(col_begin_plan & r + a & ":ZZ" & r + a).Interior.Color = xlNone
            TaakBalkPlaatsen CKalender, t.startdatum, t.einddatum, t.Status, r + a, True, ws
            t.update
        End If
    End If
    a = a + 1
Loop
 turbo_UIT
End Function

Function laatstekolom() As Long
    Dim ws As String
    ws = ActiveSheet.Name
    laatstekolom = ThisWorkbook.Sheets(ws).Range("XFD" & SoortPlanning.row_dag).End(xlToLeft).Column
End Function
