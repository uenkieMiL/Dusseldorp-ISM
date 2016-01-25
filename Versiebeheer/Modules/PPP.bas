Attribute VB_Name = "PPP"
    Option Explicit
    
    Public Const startkolom = 19
    Public Const startrij = 8
    
    Public Const col_synergy As String = "A"
    Public Const col_vestiging As String = "B"
    Public Const col_bedrijf As String = "C"
    Public Const col_naamEnOpdrachtgever As String = "D"
    Public Const col_voornaam As String = "E"
    Public Const col_pl As String = "F"
    Public Const col_wvb As String = "G"
    Public Const col_uitv As String = "H"
    Public Const col_machinist As String = "I"
    Public Const col_timmerman As String = "J"
    Public Const col_grondwerker As String = "K"
    Public Const col_sloper As String = "L"
    Public Const col_dav As String = "M"
    Public Const col_dta As String = "N"
    Public Const col_kvp As String = "O"
    Public Const col_hvk As String = "P"
    Public Const col_uitvoerder As String = "Q"
    Public Const col_beoordeling As String = "R"
    
    Public Const col_plan_start As String = "S"
    Public Const col_plan_eind As String = "TR"
    
    Public Const row_datum As Long = 1
    Public Const row_jaar As Long = 2
    Public Const row_maand As Long = 3
    Public Const row_week As Long = 4
    Public Const row_dag As Long = 5
    
    Function MaakProjectPersoneelsPlanning()
    Dim kalender As New Collection
    Dim d As datum
    Dim lk As Long
    Dim u As Uursoort
    Dim lr As Long
    Dim pp As PersoneelPlanning
    Dim vandaag As Date
    Dim mindatum As Date
    Dim maxdatum As Date
    Dim p As Personeel
    Dim pr As project
    Dim lijstpersoneel As New Collection
    Dim r As Long
    Dim fr_start As Long, fr_eind As Long
    Blad6.Select
    Turbo_AAN
        ThisWorkbook.Sheets(Blad6.Name).Cells.UnMerge
        ThisWorkbook.Sheets(Blad6.Name).Cells.Rows.ClearOutline
        If Blad6.AutoFilterMode = True Then Blad6.AutoFilterMode = False
        vandaag = Now()
        
        mindatum = DateAdd("d", 0 - Weekday(vandaag, vbMonday) - 14, Now())
        maxdatum = DateAdd("d", (104 * 7 - 1), mindatum)
        
    Set kalender = Lijsten.KalenderStartEind(mindatum, maxdatum)
    Set lijstpersoneel = LijstProjectPersoneelLijstOphalen(kalender)
    'Set lijstpersoneel = PersoneelLijstOphalen(kalender)
    lr = ThisWorkbook.Sheets(Blad6.Name).Range(col_bedrijf & "1048576").End(xlUp).Row
    lk = ThisWorkbook.Sheets(Blad6.Name).Range("XFD" & PPP.row_dag).End(xlToLeft).Column
    If lr > 8 Then ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_synergy & PPP.startrij & ":" & PPP.col_plan_eind & lr).Clear
    If lk >= PPP.startkolom Then ThisWorkbook.Sheets(Blad6.Name).Range(Range(Columns(PPP.startkolom), Columns(lk + 1)).Address).delete Shift:=xlToLeft
    If lk >= PPP.startkolom Then ThisWorkbook.Sheets(Blad6.Name).Cells.Interior.Color = xlNone
    
    PlaatsFeestdagenPPP kalender
    MaakKolomVandaagPPP kalender
    
    r = 8
    
    For Each u In lijstpersoneel
        Range(PPP.col_synergy & r & ":" & PPP.col_beoordeling & r).Merge
        Range(PPP.col_synergy & r).Value = u.Omschrijving
        Range(PPP.col_synergy & r & ":" & PPP.col_beoordeling & r).Interior.Color = u.Kleur
        Rows(r & ":" & r).RowHeight = 15
        If u.Koppelbaar = True Then
            For Each pr In u.CProjecten
                If ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_vestiging & r) <> pr.Vestiging And ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_vestiging & r) <> "" Then
                    r = r + 1
                    ThisWorkbook.Sheets(Blad6.Name).Range(r & ":" & r).Interior.Color = 1
                    Rows(r & ":" & r).RowHeight = 5
                End If
                r = r + 1
                
                Range(PPP.col_synergy & r) = pr.synergy
                Range(PPP.col_vestiging & r) = pr.Vestiging
                Range(PPP.col_bedrijf & r) = pr.Omschrijving
                Range(PPP.col_naamEnOpdrachtgever & r) = pr.Opdrachtgever
                Range(PPP.col_synergy & r & ":" & PPP.col_beoordeling & r).Interior.Color = 12566463
                Range(PPP.col_pl & r) = pr.pl
                Range(PPP.col_wvb & r) = pr.wvb
                Range(PPP.col_uitv & r) = pr.uitv
                Rows(r & ":" & r).RowHeight = 15
                If pr.AantalPersoneel <> 0 Then fr_start = r + 1
                
    
                For Each p In pr.CPersoneel
                    Rows(r & ":" & r).RowHeight = 15
                    r = r + 1
                    Range(PPP.col_synergy & r) = pr.synergy
                    Range(PPP.col_vestiging & r) = pr.Vestiging
                    Range(PPP.col_bedrijf & r) = p.Bedrijf.Bedrijfsnaam
                    Range(PPP.col_naamEnOpdrachtgever & r) = p.Achternaam
                    Range(PPP.col_voornaam & r) = p.Naam
                    Range(PPP.col_pl & r) = pr.pl
                    Range(PPP.col_wvb & r) = pr.wvb
                    Range(PPP.col_uitv & r) = pr.uitv
                    If p.Machinist = True Then Range(PPP.col_machinist & r) = "X"
                    If p.Timmerman = True Then Range(PPP.col_timmerman & r) = "X"
                    If p.Grondwerker = True Then Range(PPP.col_grondwerker & r) = "X"
                    If p.Sloper = True Then Range(PPP.col_sloper & r) = "X"
                    If p.DHV = True Then Range(PPP.col_dav & r) = "X"
                    If p.DTA = True Then Range(PPP.col_dta & r) = "X"
                    If p.KVP = True Then Range(PPP.col_kvp & r) = "X"
                    If p.HVK = True Then Range(PPP.col_hvk & r) = "X"
                    If p.Uitvoerder = True Then Range(PPP.col_uitvoerder & r) = "X"
                    Range(PPP.col_beoordeling & r) = p.Beoordeling
                    For Each pp In p.CPersoneelPlanning
                        Range(Cells(r, pp.Kolomnummer + PPP.startkolom).Address).Interior.Color = u.Kleur
                    Next pp
                     
                Next p
                If pr.AantalPersoneel <> 0 Then
                    ThisWorkbook.Sheets(Blad6.Name).Range(fr_start & ":" & r).Group
                End If
            Next pr
        Else
            If u.AantalPersoneel <> 0 Then fr_start = r + 1
            For Each p In u.CPersoneel
                r = r + 1
                Rows(r & ":" & r).RowHeight = 15
                
                Range(PPP.col_synergy & r) = u.Omschrijving
                Range(PPP.col_bedrijf & r) = p.Bedrijf.Bedrijfsnaam
                Range(PPP.col_naamEnOpdrachtgever & r) = p.Achternaam
                Range(PPP.col_voornaam & r) = p.Naam
                If p.Machinist = True Then Range(PPP.col_machinist & r) = "X"
                If p.Timmerman = True Then Range(PPP.col_timmerman & r) = "X"
                If p.Grondwerker = True Then Range(PPP.col_grondwerker & r) = "X"
                If p.Sloper = True Then Range(PPP.col_sloper & r) = "X"
                If p.DHV = True Then Range(PPP.col_dav & r) = "X"
                If p.DTA = True Then Range(PPP.col_dta & r) = "X"
                If p.KVP = True Then Range(PPP.col_kvp & r) = "X"
                If p.HVK = True Then Range(PPP.col_hvk & r) = "X"
                If p.Uitvoerder = True Then Range(PPP.col_uitvoerder & r) = "X"
                Range(PPP.col_beoordeling & r) = p.Beoordeling
                For Each pp In p.CPersoneelPlanning
                    Range(Cells(r, pp.Kolomnummer + PPP.startkolom).Address).Interior.Color = u.Kleur
                Next pp
                 
            Next p
            If u.AantalPersoneel <> 0 Then
                ThisWorkbook.Sheets(Blad6.Name).Range(fr_start & ":" & r).Group
            End If
        End If
        r = r + 2
    Next u
    
    
    lk = 104 * 5 + PPP.startkolom - 1
    MaakRaster ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_synergy & PPP.startrij & ":" & Functies.KolomNaarLetter(CInt(lk)) & r - 2)
    
    
    ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_pl & PPP.startrij & ":" & PPP.col_beoordeling & r).HorizontalAlignment = xlCenter
    ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_synergy & PPP.startrij - 1 & ":" & PPP.col_synergy & r).HorizontalAlignment = xlCenter
    ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_synergy & PPP.startrij - 1 & ":" & PPP.col_plan_eind & r).AutoFilter
    Range(PPP.col_plan_start & ":" & col_plan_eind).ColumnWidth = 3
    For Each d In kalender
        If d.Kolomnummer > -1 Then
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_datum, PPP.startkolom + d.Kolomnummer).Address) = d.datum
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_jaar, PPP.startkolom + d.Kolomnummer).Address) = Year((d.datum))
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_maand, PPP.startkolom + d.Kolomnummer).Address) = MonthName(Month(d.datum))
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_week, PPP.startkolom + d.Kolomnummer).Address) = DatePart("ww", d.datum, vbMonday, vbFirstFourDays)
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_dag, PPP.startkolom + d.Kolomnummer).Address) = Day(d.datum)
        ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_dag, PPP.startkolom + d.Kolomnummer).Address).HorizontalAlignment = xlCenter
        End If
    Next d
    
    'MaakRaster ThisWorkbook.Sheets(blad6.Name).Range("A8:JN" & 8 + lijstpersoneel.Count)
    
    DagJaarSamenvoegenProjectPersoneel
    DagMaandSamenvoegenProjectPersoneel
    DagWeekSamenvoegenProjectPersoneel
    
    
    ThisWorkbook.Sheets(Blad6.Name).Outline.ShowLevels RowLevels:=1
    Range("A1").Select
    turbo_UIT
    End Function
    
    
    Function DagJaarSamenvoegenProjectPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PPP.startkolom
    lk = ThisWorkbook.Sheets(Blad6.Name).Range("XFD" & PPP.row_jaar).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = PPP.startkolom + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_jaar, k).Address) <> ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_jaar, k - 1).Address) Then
        Set rng = ThisWorkbook.Sheets(Blad6.Name).Range(Range(Cells(PPP.row_jaar, sk1), Cells(PPP.row_jaar, k - 1)).Address)
        With rng
            .Merge
            .HorizontalAlignment = xlCenter
        End With
        
            With rng.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        sk1 = k
    End If
    
    Next k
    Application.DisplayAlerts = True
    End Function
    
    Function DagMaandSamenvoegenProjectPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PPP.startkolom
    lk = ThisWorkbook.Sheets(Blad6.Name).Range("XFD" & PPP.row_maand).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_maand, k).Address) <> ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_maand, k - 1).Address) Then
    
        Set rng = ThisWorkbook.Sheets(Blad6.Name).Range(Range(Cells(PPP.row_maand, sk1), Cells(PPP.row_maand, k - 1)).Address)
        With rng
            .Merge
            .HorizontalAlignment = xlCenter
        End With
        
            With rng.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    
            
        With ThisWorkbook.Sheets(Blad6.Name).Range(Range(Cells(PPP.row_maand, sk1), Cells(PPP.row_maand, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
    
        sk1 = k
    End If
    
    Next k
    Application.DisplayAlerts = True
    End Function
    
    
    Function DagWeekSamenvoegenProjectPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PPP.startkolom
    lk = ThisWorkbook.Sheets(Blad6.Name).Range("XFD" & PPP.row_week).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_week, k).Address) <> ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_week, k - 1).Address) Then
    
        
            With ThisWorkbook.Sheets(Blad6.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
        
            With ThisWorkbook.Sheets(Blad6.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
            
            Set rng = ThisWorkbook.Sheets(Blad6.Name).Range(Range(Cells(PPP.row_week, sk1), Cells(PPP.row_week, k - 1)).Address)
            With rng
                .Merge
                .HorizontalAlignment = xlCenter
            End With
        
            With rng.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With rng.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
            
        With ThisWorkbook.Sheets(Blad6.Name).Range(Range(Cells(PPP.row_week, sk1), Cells(PPP.row_week, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
        
        sk1 = k
      End If
    Next k
      
End Function


Function LijstProjectPersoneelLijstOphalen(CKalender As Collection) As Collection
Dim u As Uursoort
Dim pr As project
Dim p As Personeel
Dim pp As PersoneelPlanning
Dim startDatumLijst As Date
Dim lijstprojecten As Variant
Dim lijstPlanning As Variant
Dim lijstuursoort As Variant
Dim projecten As New Collection
Dim Uursoorten As New Collection
Dim huidigesynergy As String
Dim huidigepersoneelid As Long
Dim proj As Long
Dim db As New DataBase
Dim x As Long

Dim oudproject As String
Dim oudpersoneelId As Long
Dim oudpersoneelplanningid As Long

startDatumLijst = ThisWorkbook.Sheets(Blad6.Name).Range(PPP.col_plan_start & "1")
Set LijstProjectPersoneelLijstOphalen = New Collection

lijstprojecten = db.getLijstBySQL("SELECT * FROM PROJECTEN WHERE Synergy IN (SELECT DISTINCT Synergy FROM PLANNING_PERSONEEL " & _
" WHERE(((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#))) ORDER BY PROJECTEN.Vestiging, PROJECTEN.Synergy;")

lijstuursoort = db.getLijstBySQL("SELECT * FROM UURSOORT WHERE InActief = False ORDER BY Id")

If IsEmpty(lijstuursoort) = False Then
    For x = 0 To UBound(lijstuursoort, 2)
        Set u = New Uursoort
        u.Id = lijstuursoort(0, x)
        u.Omschrijving = lijstuursoort(1, x)
        u.Kleur = lijstuursoort(2, x)
        u.Koppelbaar = lijstuursoort(3, x)
        Uursoorten.Add u, CStr(u.Id)
    Next x
End If



    For Each u In Uursoorten
        If u.Koppelbaar = True Then
        lijstPlanning = db.getLijstBySQL("SELECT PLANNING_PERSONEEL.Synergy, PROJECTEN.*, PERSONEEL.*, BEDRIJVEN.*, * " & _
                        "FROM BEDRIJVEN INNER JOIN ((PLANNING_PERSONEEL INNER JOIN PROJECTEN ON PLANNING_PERSONEEL.Synergy = PROJECTEN.Synergy) INNER JOIN PERSONEEL ON PLANNING_PERSONEEL.PersoneelId = PERSONEEL.Id) ON BEDRIJVEN.Id = PERSONEEL.BedrijfId " & _
                        "WHERE ((((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#)) AND (PLANNING_PERSONEEL.UursoortId) = " & u.Id & ") " & _
                        "ORDER BY PROJECTEN.Vestiging, PLANNING_PERSONEEL.Synergy, PLANNING_PERSONEEL.PersoneelId, PLANNING_PERSONEEL.Datum;")
        Else
        lijstPlanning = db.getLijstBySQL("SELECT PERSONEEL.*, BEDRIJVEN.*, PLANNING_PERSONEEL.*, * " & _
                        "FROM PLANNING_PERSONEEL INNER JOIN (BEDRIJVEN INNER JOIN PERSONEEL ON BEDRIJVEN.Id = PERSONEEL.BedrijfId) ON PLANNING_PERSONEEL.PersoneelId = PERSONEEL.Id " & _
                        "WHERE ((((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#)) AND (PLANNING_PERSONEEL.UursoortId) = " & u.Id & ") " & _
                        "ORDER BY PERSONEEL.Id, PLANNING_PERSONEEL.Datum;")
        End If
        
        If IsEmpty(lijstPlanning) = False Then
            If u.Koppelbaar = True Then
            'uren gekoppeld aan project
                For x = 0 To UBound(lijstPlanning, 2)
                    If x = 0 Then
                        Set pr = New project
                        Set pr = PPP.LijstNaarProject(lijstPlanning, x)
                        
                        
                        Set p = New Personeel
                        Set p = PPP.lijstNaarPersoneel(lijstPlanning, x)
                                                
                        Set pp = New PersoneelPlanning
                        Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                        p.PersoneelPlanningenToevoegen pp
                        
                    ElseIf x = UBound(lijstPlanning, 2) Then
                        If pr.synergy = lijstPlanning(1, x) Then
                        'project is gelijk
                            If p.Id = lijstPlanning(39, x) Then
                                'project en personeel is gelijk, voeg alleen de uren toe aan het huidige persoon
                                Set pp = New PersoneelPlanning
                                Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                                'voeg toe aan personeel
                                p.PersoneelPlanningenToevoegen pp
                                pr.ToevoegenPersoneel p
                                u.ToevoegenProject pr
                                
                            Else
                                pr.ToevoegenPersoneel p
                                Set p = New Personeel
                                Set p = PPP.lijstNaarPersoneel(lijstPlanning, x)
                                
                                
                                Set pp = New PersoneelPlanning
                                Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                                
                                'voeg toe aan personeel
                                p.PersoneelPlanningenToevoegen pp
                                pr.ToevoegenPersoneel p
                                u.ToevoegenProject pr
                            End If
                        Else
                            u.ToevoegenProject pr
                            
                            Set pr = New project
                            Set pr = PPP.LijstNaarProject(lijstPlanning, x)
                            
                            Set p = New Personeel
                            Set p = PPP.lijstNaarPersoneel(lijstPlanning, x)
                            
                            
                            Set pp = New PersoneelPlanning
                            Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                            pr.ToevoegenPersoneel p
                            u.ToevoegenProject pr
                        End If
                    Else
                    'tussen regel 1 en laatste regel
                        If pr.synergy = lijstPlanning(1, x) Then
                        'project is gelijk
                            If p.Id = lijstPlanning(41, x) Then
                                'project en personeel is gelijk, voeg alleen de uren toe aan het huidige persoon
                                Set pp = New PersoneelPlanning
                                Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                                'voeg toe aan personeel
                                p.PersoneelPlanningenToevoegen pp
                            Else
                                pr.ToevoegenPersoneel p
                                Set p = New Personeel
                                Set p = PPP.lijstNaarPersoneel(lijstPlanning, x)
                                
                                Set pp = New PersoneelPlanning
                                Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                                
                                'voeg toe aan personeel
                                p.PersoneelPlanningenToevoegen pp
                            End If
                        Else
                            pr.ToevoegenPersoneel p
                            u.ToevoegenProject pr
                            
                            Set pr = New project
                            Set pr = PPP.LijstNaarProject(lijstPlanning, x)
                            
                            Set p = New Personeel
                            Set p = PPP.lijstNaarPersoneel(lijstPlanning, x)
                            
                            Set pp = New PersoneelPlanning
                            Set pp = PPP.LijstNaarPersoneelPlanning(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                        End If
                    End If
                Next x
            Else
                For x = 0 To UBound(lijstPlanning, 2)
                'Uren zijn niet koppelbaar
                    If x = 0 Then
                        'eerste regel
                        Set p = New Personeel
                        Set p = LijstNaarPersoneelNietKoppelbaar(lijstPlanning, x)
                        
                        Set pp = New PersoneelPlanning
                        Set pp = LijstNaarPersoneelPlanningNietKoppelbaar(lijstPlanning, x, CKalender)
                        
                        p.PersoneelPlanningenToevoegen pp
                        
                    ElseIf x = UBound(lijstPlanning, 2) Then
                        If p.Id = lijstPlanning(0, x) Then
                            'persoon is gelijk
                            Set pp = New PersoneelPlanning
                            Set pp = LijstNaarPersoneelPlanningNietKoppelbaar(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                            u.ToevoegenPersoneel p
                        Else
                            u.ToevoegenPersoneel p
                            Set p = New Personeel
                            Set p = LijstNaarPersoneelNietKoppelbaar(lijstPlanning, x)
                            
                            Set pp = New PersoneelPlanning
                            Set pp = LijstNaarPersoneelPlanningNietKoppelbaar(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                            u.ToevoegenPersoneel p
                        End If
                    Else
                       If p.Id = lijstPlanning(0, x) Then
                            'persoon is gelijk
                            Set pp = New PersoneelPlanning
                            Set pp = LijstNaarPersoneelPlanningNietKoppelbaar(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                        Else
                            u.ToevoegenPersoneel p
                            Set p = New Personeel
                            Set p = LijstNaarPersoneelNietKoppelbaar(lijstPlanning, x)
                            
                            Set pp = New PersoneelPlanning
                            Set pp = LijstNaarPersoneelPlanningNietKoppelbaar(lijstPlanning, x, CKalender)
                            
                            p.PersoneelPlanningenToevoegen pp
                        End If
                    End If
                
                Next x
            End If
        End If
    LijstProjectPersoneelLijstOphalen.Add u
    
    Next u



End Function


Function LijstNaarProject(lijst As Variant, r As Long) As project
        Dim pr As New project
        pr.synergy = lijst(0, r)
        pr.Omschrijving = lijst(2, r)
        pr.Opdrachtgever = lijst(3, r)
        pr.pv = lijst(4, r)
        pr.pl = lijst(5, r)
        pr.CALC = lijst(6, r)
        pr.wvb = lijst(7, r)
        pr.uitv = lijst(8, r)
        pr.Vestiging = lijst(11, r)
        Set LijstNaarProject = pr
End Function

Function lijstNaarPersoneel(lijst As Variant, r As Long) As Personeel
    Dim p As New Personeel
    p.Id = lijst(17, r)
    p.Achternaam = lijst(18, r)
    p.Naam = lijst(19, r)
    p.BSN = lijst(20, r)
    p.Machinist = lijst(21, r)
    p.Timmerman = lijst(22, r)
    p.Grondwerker = lijst(23, r)
    p.Sloper = lijst(24, r)
    p.DHV = lijst(25, r)
    p.DTA = lijst(26, r)
    p.Uitvoerder = lijst(27, r)
    p.Bijzonderheden = lijst(28, r)
    p.Beoordeling = lijst(29, r)
    p.BedrijfId = lijst(31, r)
    p.KVP = lijst(32, r)
    p.HVK = lijst(33, r)
    p.Bedrijf.Id = lijst(34, r)
    p.Bedrijf.Bedrijfsnaam = lijst(36, r)
    Set lijstNaarPersoneel = p
End Function

Function LijstNaarPersoneelPlanning(lijst As Variant, r As Long, ByRef CKalender As Collection) As PersoneelPlanning
    Dim pp As New PersoneelPlanning
    pp.personeelid = lijst(41, r)
    pp.datum = lijst(42, r)
    pp.UursoortId = lijst(43, r)
    pp.Kolomnummer = HaalKolomnummerOp(CKalender, pp.datum)
    Set LijstNaarPersoneelPlanning = pp
End Function

Function LijstNaarPersoneelPlanningNietKoppelbaar(lijst As Variant, r As Long, ByRef CKalender As Collection) As PersoneelPlanning
    Dim pp As New PersoneelPlanning
    pp.personeelid = lijst(24, r)
    pp.datum = lijst(25, r)
    pp.UursoortId = lijst(26, r)
    pp.Kolomnummer = HaalKolomnummerOp(CKalender, pp.datum)
    Set LijstNaarPersoneelPlanningNietKoppelbaar = pp
End Function

Function LijstNaarPersoneelNietKoppelbaar(lijst As Variant, r As Long) As Personeel
    Dim p As New Personeel
    p.Id = lijst(0, r)
    p.Achternaam = lijst(1, r)
    p.Naam = lijst(2, r)
    p.BSN = lijst(3, r)
    p.Machinist = lijst(4, r)
    p.Timmerman = lijst(5, r)
    p.Grondwerker = lijst(6, r)
    p.Sloper = lijst(7, r)
    p.DHV = lijst(8, r)
    p.DTA = lijst(9, r)
    p.Uitvoerder = lijst(10, r)
    p.Bijzonderheden = lijst(11, r)
    p.Beoordeling = lijst(12, r)
    p.BedrijfId = lijst(14, r)
    p.DTA = lijst(15, r)
    p.DTA = lijst(16, r)
    p.Bedrijf.Id = lijst(17, r)
    p.Bedrijf.Bedrijfsnaam = lijst(19, r)
    Set LijstNaarPersoneelNietKoppelbaar = p
End Function

Function PlaatsFeestdagenPPP(CKalender As Collection)
Dim d As datum
Dim rng As Range
For Each d In CKalender
        If d.Kolomnummer > -1 And d.feestdag = True Then
            Set rng = ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_dag, PPP.startkolom + d.Kolomnummer), Cells(laatsterij, PPP.startkolom + d.Kolomnummer))
            rng.Interior.Color = kleuren.feestdag
        End If
    Next d
End Function
Function MaakKolomVandaagPPP(CKalender As Collection)
Dim vandaag As Date: vandaag = Now(): vandaag = FormatDateTime(Date, vbShortDate)
Dim d As datum
Dim gevonden As Boolean
Dim kolom As Long
Dim rng As Range
Do While gevonden = False
    For Each d In CKalender
        If d.datum = vandaag Then
            kolom = d.Kolomnummer
            gevonden = True
            Exit For
        End If
    Next d
    
    If gevonden = False Then vandaag = DateAdd("d", 1, vandaag)
Loop

Set rng = ThisWorkbook.Sheets(Blad6.Name).Range(Cells(PPP.row_dag, PPP.startkolom + kolom), Cells(laatsterij, PPP.startkolom + kolom))
rng.Interior.Color = kleuren.vandaag
End Function

Function laatsterij() As Long
    laatsterij = ThisWorkbook.Sheets(Blad5.Name).Range(PPP.col_bedrijf & 1048576).End(xlUp).Row
End Function
