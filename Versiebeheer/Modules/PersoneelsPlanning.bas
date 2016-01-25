Attribute VB_Name = "PersoneelsPlanning"
    Public Const startkolom As Integer = 17
    Public Const startrij As Integer = 8
    Public Const col_pers_id As String = "A"
    Public Const col_pers_bedrijf As String = "B"
    Public Const col_pers_achternaam As String = "C"
    Public Const col_pers_voornaam As String = "D"
    Public Const col_pers_BSN As String = "E"
    Public Const col_pers_machinist As String = "F"
    Public Const col_pers_timmerman As String = "G"
    Public Const col_pers_grondwerker As String = "H"
    Public Const col_pers_sloper As String = "I"
    Public Const col_pers_DAV As String = "J"
    Public Const col_pers_DTA As String = "K"
    Public Const col_pers_KVP As String = "L"
    Public Const col_pers_HVK As String = "M"
    Public Const col_pers_UITV As String = "N"
    Public Const col_pers_beoordeling As String = "O"
    Public Const col_pers_bijz As String = "P"
    Public Const col_plan_start As String = "Q"
    Public Const col_plan_eind As String = "TP"
    Public Const row_datum As Integer = 1
    Public Const row_jaar As Integer = 2
    Public Const row_maand As Integer = 3
    Public Const row_week As Integer = 4
    Public Const row_dag As Integer = 5
    
    
    Function MaakPersoneelsPlanning()
    Dim kalender As New Collection
    Dim d As datum
    Dim lk As Long
    Dim lr As Long
    Dim pp As PersoneelPlanning
    Dim vandaag As Date
    Dim mindatum As Date
    Dim maxdatum As Date
    Dim p As Personeel
    Dim lijstpersoneelintern As New Collection
    Dim lijstpersoneelextern As New Collection
    Dim lijstprojecen As New Collection
    Dim project As project
    
    Dim r As Long
    Turbo_AAN
        If Blad5.Name <> ActiveSheet.Name Then Blad5.Select
        Range(PersoneelsPlanning.col_pers_id & PersoneelsPlanning.startrij).Select
        
        If Blad5.AutoFilterMode = True Then Blad5.AutoFilterMode = False
        vandaag = Now()
        
        mindatum = DateAdd("d", 0 - Weekday(vandaag, vbMonday) - 14, Now())
        maxdatum = DateAdd("d", (104 * 7 - 1), mindatum)
        
        
    ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_bijz & "1") = mindatum + 1
    Set kalender = Lijsten.KalenderStartEind(mindatum, maxdatum)
    Set lijstprojecen = lijstprojectenophalen
    Set lijstpersoneelintern = PersoneelLijstOphalenIntern(kalender)
    Set lijstpersoneelextern = PersoneelLijstOphalenExtern(kalender)
    lr = ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_achternaam & 1048576).End(xlUp).Row
    lk = ThisWorkbook.Sheets(Blad5.Name).Range("XFD" & row_dag).End(xlToLeft).Column
    If lr > PersoneelsPlanning.startrij Then ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_id & PersoneelsPlanning.startrij & ":" & PersoneelsPlanning.col_plan_eind & lr).Clear
    If lk >= PersoneelsPlanning.startkolom Then ThisWorkbook.Sheets(Blad5.Name).Range(Range(Columns(PersoneelsPlanning.startkolom), Columns(lk + 1)).Address).delete Shift:=xlToLeft
    
    ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.startrij, PersoneelsPlanning.startkolom), Cells(lr, 16384)).Interior.Color = xlNone
    
    For Each d In kalender
        If d.Kolomnummer > -1 Then
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_datum, PersoneelsPlanning.startkolom + d.Kolomnummer).Address) = d.datum
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_jaar, PersoneelsPlanning.startkolom + d.Kolomnummer).Address) = Year((d.datum))
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_maand, PersoneelsPlanning.startkolom + d.Kolomnummer).Address) = MonthName(Month(d.datum))
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_week, PersoneelsPlanning.startkolom + d.Kolomnummer).Address) = DatePart("ww", d.datum, vbMonday, vbFirstFourDays)
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_dag, PersoneelsPlanning.startkolom + d.Kolomnummer).Address) = Day(d.datum)
        ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_dag, PersoneelsPlanning.startkolom + d.Kolomnummer).Address).HorizontalAlignment = xlCenter
        End If
    Next d
    lk = ThisWorkbook.Sheets(Blad5.Name).Range("XFD" & row_dag).End(xlToLeft).Column
   
    MaakRaster ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_id & PersoneelsPlanning.startrij & ":" & KolomNaarLetter(CInt(lk)) & PersoneelsPlanning.startrij + lijstpersoneelintern.Count + lijstpersoneelextern.Count + 5)
    
    DagJaarSamenvoegenPersoneel
    DagMaandSamenvoegenPersoneel
    DagWeekSamenvoegenPersoneel
    
    PersoneelsPlanning.PlaatsFeestdagen kalender, lijstpersoneelextern.Count + lijstpersoneelintern.Count
    PersoneelsPlanning.MaakKolomVandaag kalender, lijstpersoneelextern.Count + lijstpersoneelintern.Count
    r = 8
    
    
    For Each p In lijstpersoneelintern
        If oudebedrijfsnaam <> p.Bedrijf.Bedrijfsnaam And oudebedrijfsnaam <> "" Then
         ThisWorkbook.Sheets(Blad5.Name).Range(r & ":" & r).Interior.Color = 1
        Rows(r & ":" & r).RowHeight = 5
        r = r + 1
        
        End If
        ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_id & r) = p.Id
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bedrijf & r) = p.Bedrijf.Bedrijfsnaam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_achternaam & r) = p.Achternaam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_voornaam & r) = p.Naam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_BSN & r) = p.BSN
        If p.Machinist Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_machinist & r) = "X"
        If p.Timmerman Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_timmerman & r) = "X"
        If p.Grondwerker Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_grondwerker & r) = "X"
        If p.Sloper Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_sloper & r) = "X"
        If p.DHV Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_DAV & r) = "X"
        If p.DTA Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_DTA & r) = "X"
        If p.KVP Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_KVP & r) = "X"
        If p.HVK Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_HVK & r) = "X"
        If p.Uitvoerder Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_UITV & r) = "X"
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_beoordeling & r) = p.Beoordeling
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bijz & r) = p.Bijzonderheden
        
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bijz & r).HorizontalAlignment = xlLeft
        If p.AantalPersoneelPlanningen > 0 Then
            For Each pp In p.CPersoneelPlanning
                 Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address).Interior.Color = pp.Uursoort.Kleur
                If Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = "" Then
                    If pp.synergy = "" Then
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = UCase(Left(pp.Uursoort.Omschrijving, 5))
                    Else
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = pp.synergy
                    End If
                Else
                    If pp.synergy = "" Then
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = _
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) & Chr(10) & Chr(13) & UCase(Left(pp.Uursoort.Omschrijving, 5))
                    Else
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = _
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) & Chr(10) & Chr(13) & pp.synergy
                    End If
                End If
                Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address).HorizontalAlignment = xlCenter
                     
            Next pp
        End If
        Rows(r & ":" & r).EntireRow.AutoFit
        r = r + 1
        oudebedrijfsnaam = p.Bedrijf.Bedrijfsnaam
    Next p
        
        ThisWorkbook.Sheets(Blad5.Name).Range(r & ":" & r).Interior.Color = 1
        Rows(r & ":" & r).RowHeight = 5
        r = r + 1
    For Each p In lijstpersoneelextern
        ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_id & r) = p.Id
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bedrijf & r) = p.Bedrijf.Bedrijfsnaam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_achternaam & r) = p.Achternaam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_voornaam & r) = p.Naam
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_BSN & r) = p.BSN
        If p.Machinist Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_machinist & r) = "X"
        If p.Timmerman Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_timmerman & r) = "X"
        If p.Grondwerker Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_grondwerker & r) = "X"
        If p.Sloper Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_sloper & r) = "X"
        If p.DHV Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_DAV & r) = "X"
        If p.DTA Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_DTA & r) = "X"
        If p.KVP Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_KVP & r) = "X"
        If p.HVK Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_HVK & r) = "X"
        If p.Uitvoerder Then ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_UITV & r) = "X"
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_beoordeling & r) = p.Beoordeling
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bijz & r) = p.Bijzonderheden
        
        ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_bijz & r).HorizontalAlignment = xlLeft
        If p.AantalPersoneelPlanningen > 0 Then
            For Each pp In p.CPersoneelPlanning
                Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address).Interior.Color = pp.Uursoort.Kleur
                If Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = "" Then
                    If pp.synergy = "" Then
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = UCase(Left(pp.Uursoort.Omschrijving, 5))
                    Else
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = pp.synergy
                    End If
                Else
                    If pp.synergy = "" Then
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = _
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) & Chr(10) & Chr(13) & UCase(Left(pp.Uursoort.Omschrijving, 5))
                    Else
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) = _
                        Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address) & Chr(10) & Chr(13) & pp.synergy
                    End If
                End If
                Range(Range(Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer), Cells(r, PersoneelsPlanning.startkolom + pp.Kolomnummer)).Address).HorizontalAlignment = xlCenter
                     
            Next pp
        End If
        
        Rows(r & ":" & r).EntireRow.AutoFit
        r = r + 1
    Next p
    
    ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_machinist & "8:" & PersoneelsPlanning.col_pers_beoordeling & r).HorizontalAlignment = xlCenter
    
    ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_id & PersoneelsPlanning.startrij - 1 & ":" & PersoneelsPlanning.col_plan_eind & r).AutoFilter
    
    MaakVoorwardelijkeOpmaak ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_achternaam & PersoneelsPlanning.startrij & ":" & PersoneelsPlanning.col_pers_voornaam & r)
    Range("A1").Select
    turbo_UIT
    Set kalender = Nothing
    Set lijstpersoneelintern = Nothing
    Set lijstpersoneelextern = Nothing
    End Function
    

    Function DagJaarSamenvoegenPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PersoneelsPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad5.Name).Range("XFD" & PersoneelsPlanning.row_jaar).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = PersoneelsPlanning.startkolom + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_jaar, k).Address) <> ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_jaar, k - 1).Address) Then
        Set rng = ThisWorkbook.Sheets(Blad5.Name).Range(Range(Cells(PersoneelsPlanning.row_jaar, sk1), Cells(PersoneelsPlanning.row_jaar, k - 1)).Address)
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
    
    Function DagMaandSamenvoegenPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PersoneelsPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad5.Name).Range("XFD" & PersoneelsPlanning.row_maand).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_maand, k).Address) <> ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_maand, k - 1).Address) Then
    
        Set rng = ThisWorkbook.Sheets(Blad5.Name).Range(Range(Cells(PersoneelsPlanning.row_maand, sk1), Cells(PersoneelsPlanning.row_maand, k - 1)).Address)
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
    
            
        With ThisWorkbook.Sheets(Blad5.Name).Range(Range(Cells(PersoneelsPlanning.row_maand, sk1), Cells(PersoneelsPlanning.row_maand, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
    
        sk1 = k
    End If
    
    Next k
    Application.DisplayAlerts = True
    End Function
    
    
    Function DagWeekSamenvoegenPersoneel()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = PersoneelsPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad5.Name).Range("XFD" & PersoneelsPlanning.row_week).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_week, k).Address) <> ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_week, k - 1).Address) Then
    
        
            With ThisWorkbook.Sheets(Blad5.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
        
            With ThisWorkbook.Sheets(Blad5.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
            
            Set rng = ThisWorkbook.Sheets(Blad5.Name).Range(Range(Cells(PersoneelsPlanning.row_week, sk1), Cells(PersoneelsPlanning.row_week, k - 1)).Address)
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
            
        With ThisWorkbook.Sheets(Blad5.Name).Range(Range(Cells(PersoneelsPlanning.row_week, sk1), Cells(PersoneelsPlanning.row_week, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
        
        sk1 = k
      End If
    Next k
      
End Function


Function PersoneelLijstOphalenIntern(CKalender As Collection) As Collection
Dim lijst As Variant
Dim lijstPlanning As Variant
Dim lijstPersoneelPlanning As New Collection
Dim p As Personeel
Dim startDatumLijst As Date
Dim pp As PersoneelPlanning
startDatumLijst = ThisWorkbook.Sheets(Blad5.Name).Range("P1")
Set PersoneelLijstOphalenIntern = New Collection
Dim db As New DataBase
lijst = db.getLijstBySQL("SELECT PERSONEEL.*, BEDRIJVEN.* FROM BEDRIJVEN INNER JOIN PERSONEEL ON BEDRIJVEN.Id = PERSONEEL.BedrijfId WHERE PERSONEEL.BedrijfId IN (1,56,58) AND PERSONEEL.ARCHIEF = False ORDER BY PERSONEEL.BedrijfId, PERSONEEL.Achternaam;")

lijstPlanning = db.getLijstBySQL("SELECT PLANNING_PERSONEEL.*, UURSOORT.*, PLANNING_PERSONEEL.Datum, * " & _
"FROM UURSOORT INNER JOIN PLANNING_PERSONEEL ON UURSOORT.Id = PLANNING_PERSONEEL.UursoortId " & _
"WHERE (((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#));")

If IsEmpty(lijstPlanning) = False Then
    For x = 0 To UBound(lijstPlanning, 2)
        If lijstPlanning(9, x) = False Then
            Set pp = New PersoneelPlanning
            pp.Id = lijstPlanning(0, x)
            pp.personeelid = lijstPlanning(1, x)
            pp.datum = lijstPlanning(2, x)
            pp.Kolomnummer = HaalKolomnummerOp(CKalender, pp.datum)
            pp.UursoortId = lijstPlanning(3, x)
            pp.synergy = lijstPlanning(4, x)
            pp.Uursoort.Id = lijstPlanning(5, x)
            pp.Uursoort.Omschrijving = lijstPlanning(6, x)
            pp.Uursoort.Kleur = lijstPlanning(7, x)
            pp.Uursoort.Koppelbaar = lijstPlanning(8, x)
            lijstPersoneelPlanning.Add pp, CStr(pp.Id)
        End If
    Next x
End If

If IsEmpty(lijst) = False Then
    For x = 0 To UBound(lijst, 2)
        Set p = New Personeel
        p.Id = lijst(0, x)
        p.Achternaam = lijst(1, x)
        p.Naam = lijst(2, x)
        p.BSN = lijst(3, x)
        p.Machinist = lijst(4, x)
        p.Timmerman = lijst(5, x)
        p.Grondwerker = lijst(6, x)
        p.Sloper = lijst(7, x)
        p.DHV = lijst(8, x)
        p.DTA = lijst(9, x)
        p.Uitvoerder = lijst(10, x)
        p.Bijzonderheden = lijst(11, x)
        p.Beoordeling = lijst(12, x)
        p.Archief = lijst(13, x)
        p.BedrijfId = lijst(14, x)
        p.KVP = lijst(15, x)
        p.HVK = lijst(16, x)
        p.Bedrijf.Id = lijst(17, x)
        p.Bedrijf.KVK = lijst(18, x)
        p.Bedrijf.Bedrijfsnaam = lijst(19, x)
        If IsNull(lijst(20, x)) = False Then p.Bedrijf.Contactpersoon = lijst(20, x)
        If IsNull(lijst(21, x)) = False Then p.Bedrijf.Telefoonnummer = lijst(21, x)
        If IsNull(lijst(22, x)) = False Then p.Bedrijf.Emailadres = lijst(22, x)
        
        For Each pp In lijstPersoneelPlanning
            If pp.personeelid = p.Id Then
            p.PersoneelPlanningenToevoegen pp
            End If
        Next pp
        PersoneelLijstOphalenIntern.Add p, CStr(p.Id)
    Next x

End If


End Function

Function PersoneelLijstOphalenExtern(CKalender As Collection) As Collection
Dim lijst As Variant
Dim lijstPlanning As Variant
Dim lijstPersoneelPlanning As New Collection
Dim p As Personeel
Dim startDatumLijst As Date
Dim pp As PersoneelPlanning
Dim db As New DataBase

startDatumLijst = ThisWorkbook.Sheets(Blad5.Name).Range("P1")
Set PersoneelLijstOphalenExtern = New Collection

lijst = db.getLijstBySQL("SELECT PERSONEEL.*, BEDRIJVEN.* FROM BEDRIJVEN INNER JOIN PERSONEEL ON BEDRIJVEN.Id = PERSONEEL.BedrijfId WHERE PERSONEEL.BedrijfId NOT IN (1,56,58) AND PERSONEEL.ARCHIEF = False ORDER BY BEDRIJVEN.Bedrijfsnaam, PERSONEEL.Achternaam;")

lijstPlanning = db.getLijstBySQL("SELECT PLANNING_PERSONEEL.*, UURSOORT.*, PLANNING_PERSONEEL.Datum " & _
"FROM UURSOORT INNER JOIN PLANNING_PERSONEEL ON UURSOORT.Id = PLANNING_PERSONEEL.UursoortId " & _
"WHERE (((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#));")

If IsEmpty(lijstPlanning) = False Then
    For x = 0 To UBound(lijstPlanning, 2)
        If lijstPlanning(9, x) = False Then
            Set pp = New PersoneelPlanning
            pp.Id = lijstPlanning(0, x)
            pp.personeelid = lijstPlanning(1, x)
            pp.datum = lijstPlanning(2, x)
            pp.Kolomnummer = HaalKolomnummerOp(CKalender, pp.datum)
            pp.UursoortId = lijstPlanning(3, x)
            pp.synergy = lijstPlanning(4, x)
            pp.Uursoort.Id = lijstPlanning(5, x)
            pp.Uursoort.Omschrijving = lijstPlanning(6, x)
            pp.Uursoort.Kleur = lijstPlanning(7, x)
            pp.Uursoort.Koppelbaar = lijstPlanning(8, x)
            lijstPersoneelPlanning.Add pp, CStr(pp.Id)
        End If
    Next x
End If

If IsEmpty(lijst) = False Then
    For x = 0 To UBound(lijst, 2)
        Set p = New Personeel
        p.Id = lijst(0, x)
        p.Achternaam = lijst(1, x)
        p.Naam = lijst(2, x)
        p.BSN = lijst(3, x)
        p.Machinist = lijst(4, x)
        p.Timmerman = lijst(5, x)
        p.Grondwerker = lijst(6, x)
        p.Sloper = lijst(7, x)
        p.DHV = lijst(8, x)
        p.DTA = lijst(9, x)
        p.Uitvoerder = lijst(10, x)
        p.Bijzonderheden = lijst(11, x)
        p.Beoordeling = lijst(12, x)
        p.Archief = lijst(13, x)
        p.BedrijfId = lijst(14, x)
        p.KVP = lijst(15, x)
        p.HVK = lijst(16, x)
        p.Bedrijf.Id = lijst(17, x)
        p.Bedrijf.KVK = lijst(18, x)
        p.Bedrijf.Bedrijfsnaam = lijst(19, x)
        If IsNull(lijst(20, x)) = False Then p.Bedrijf.Contactpersoon = lijst(20, x)
        If IsNull(lijst(21, x)) = False Then p.Bedrijf.Telefoonnummer = lijst(21, x)
        If IsNull(lijst(22, x)) = False Then p.Bedrijf.Emailadres = lijst(22, x)
        
        For Each pp In lijstPersoneelPlanning
            If pp.personeelid = p.Id Then
            p.PersoneelPlanningenToevoegen pp
            End If
        Next pp
        PersoneelLijstOphalenExtern.Add p, CStr(p.Id)
    Next x

End If


End Function
Public Function KolomNaarLetter(Column As Long) As String
    If Column < 1 Then Exit Function
    KolomNaarLetter = KolomNaarLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function
Function HaalKolomnummerOp(CKalender As Collection, datum As Date) As Long
Dim d As datum

For Each d In CKalender
        If datum = d.datum Then
        HaalKolomnummerOp = d.Kolomnummer
        Exit Function
    End If
Next d
End Function

Function DeletePP(Target As Range)
Dim c As Range
Dim pp As PersoneelPlanning
For Each c In Target
If Not Intersect(c, Range(PersoneelsPlanning.col_plan_start & PersoneelsPlanning.startrij - 1 & ":" & PersoneelsPlanning.col_plan_eind & PersoneelsPlanning.laatsterij)) Is Nothing Then
    If IsNumeric(ThisWorkbook.Sheets(Blad5.Name).Range("A" & c.Row)) = True And c.Interior.Color <> xlNone Then
        Set pp = New PersoneelPlanning
        pp.datum = Cells(1, c.Column)
        pp.personeelid = Cells(c.Row, 1)
        pp.DeleteDatumPersoneel
        c.Interior.Color = xlNone
        c.ClearContents
    End If
End If

Next c

End Function


Function PlaatsFeestdagen(CKalender As Collection, Aantal As Long)
Dim d As datum
For Each d In CKalender
        If d.Kolomnummer > -1 And d.feestdag = True Then
            Set rng = ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_dag, PersoneelsPlanning.startkolom + d.Kolomnummer), Cells(PersoneelsPlanning.startrij + Aantal + 5, startkolom + d.Kolomnummer))
            rng.Interior.Color = kleuren.feestdag
        End If
    Next d
End Function
Function MaakKolomVandaag(CKalender As Collection, Aantal As Long)
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

Set rng = ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_dag, PersoneelsPlanning.startkolom + kolom), Cells(PersoneelsPlanning.startrij + Aantal + 5, PersoneelsPlanning.startkolom + kolom))
rng.Interior.Color = kleuren.vandaag
End Function


Sub MaakVoorwardelijkeOpmaak(rng As Range)
'
' Macro4 Macro
'

Blad5.Cells.FormatConditions.delete

rng.FormatConditions.Add Type:=xlExpression, Formula1:="=($" & PersoneelsPlanning.col_pers_beoordeling & PersoneelsPlanning.startrij & "=1)"

    With rng.FormatConditions(rng.FormatConditions.Count)
        .SetFirstPriority
        With .Interior
            .PatternColorIndex = xlAutomatic
            .Color = kleuren.beoordeling_goed
            .TintAndShade = 0
        End With
        StopIfTrue = False
    End With
    
    rng.FormatConditions.Add Type:=xlExpression, Formula1:="=($" & PersoneelsPlanning.col_pers_beoordeling & PersoneelsPlanning.startrij & "=2)"

    With rng.FormatConditions(rng.FormatConditions.Count)
        .SetFirstPriority
        With .Interior
            .PatternColorIndex = xlAutomatic
            .Color = kleuren.beoordeling_redelijk
            .TintAndShade = 0
        End With
        StopIfTrue = False
    End With

    rng.FormatConditions.Add Type:=xlExpression, Formula1:="=($" & PersoneelsPlanning.col_pers_beoordeling & PersoneelsPlanning.startrij & "=4)"
    With rng.FormatConditions(rng.FormatConditions.Count)
        .SetFirstPriority
        With .Interior
            .PatternColorIndex = xlAutomatic
            .Color = kleuren.beoordeling_onvoldoende
            .TintAndShade = 0
        End With
        StopIfTrue = False
    End With


End Sub

Function LijstProjectPersoneelLijstOphalen(CKalender As Collection) As Collection
Dim pr As project
Dim p As Personeel
Dim pp As PersoneelPlanning
'Dim LijstProjectPersoneelLijstOphalen As New Collection
Dim startDatumLijst As Date
Dim lijstprojecten As Variant
Dim lijstPlanning As Variant

startDatumLijst = ThisWorkbook.Sheets(Blad5.Name).Range("O1")
Set LijstProjectPersoneelLijstOphalen = New Collection

lijstprojecten = DataBase.LijstOpBasisVanQuery("SELECT * FROM PROECTEN WHERE Synergy IN (SELECT DISTINCT Synergy from PLANNING_PERSONEEL " & _
"WHERE (((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#))")

lijstPlanning = DataBase.LijstOpBasisVanQuery("SELECT UURSOORT.*, PLANNING_PERSONEEL.*, PERSONEEL.*, BEDRIJVEN.* " & _
"FROM BEDRIJVEN INNER JOIN (UURSOORT INNER JOIN (PERSONEEL INNER JOIN PLANNING_PERSONEEL ON PERSONEEL.Id = PLANNING_PERSONEEL.PersoneelId) ON UURSOORT.Id = PLANNING_PERSONEEL.UursoortId) ON BEDRIJVEN.Id = PERSONEEL.BedrijfId " & _
"WHERE (((PLANNING_PERSONEEL.Datum) >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "#)) " & _
"ORDER BY UURSOORT.Id, PLANNING_PERSONEEL.Synergy, BEDRIJVEN.Id, PERSONEEL.Achternaam, PLANNING_PERSONEEL.Datum;")



If IsEmpty(lijstPlanning) = False Then
    For x = 0 To UBound(lijstPlanning, 2)
        If lijstPlanning(9, x) = False Then
            Set pp = New PersoneelPlanning
            pp.Id = lijstPlanning(0, x)
            pp.personeelid = lijstPlanning(1, x)
            pp.datum = lijstPlanning(2, x)
            pp.Kolomnummer = HaalKolomnummerOp(CKalender, pp.datum)
            pp.UursoortId = lijstPlanning(3, x)
            pp.synergy = lijstPlanning(4, x)
            pp.Uursoort.Id = lijstPlanning(5, x)
            pp.Uursoort.Omschrijving = lijstPlanning(6, x)
            pp.Uursoort.Kleur = lijstPlanning(7, x)
            pp.Uursoort.Koppelbaar = lijstPlanning(8, x)
            lijstPersoneelPlanning.Add pp, CStr(pp.Id)
        End If
    Next x
End If

If IsEmpty(lijst) = False Then
    For x = 0 To UBound(lijst, 2)
        Set p = New Personeel
        p.Id = lijst(0, x)
        p.Achternaam = lijst(1, x)
        p.Naam = lijst(2, x)
        p.BSN = lijst(3, x)
        p.Machinist = lijst(4, x)
        p.Timmerman = lijst(5, x)
        p.Grondwerker = lijst(6, x)
        p.Sloper = lijst(7, x)
        p.DHV = lijst(8, x)
        p.DTA = lijst(9, x)
        p.Uitvoerder = lijst(10, x)
        p.Bijzonderheden = lijst(11, x)
        p.Beoordeling = lijst(12, x)
        p.Archief = lijst(13, x)
        p.BedrijfId = lijst(14, x)
        p.Bedrijf.Id = lijst(15, x)
        p.Bedrijf.KVK = lijst(16, x)
        p.Bedrijf.Bedrijfsnaam = lijst(17, x)
        If IsNull(lijst(18, x)) = False Then p.Bedrijf.Contactpersoon = lijst(18, x)
        If IsNull(lijst(19, x)) = False Then p.Bedrijf.Telefoonnummer = lijst(19, x)
        If IsNull(lijst(20, x)) = False Then p.Bedrijf.Emailadres = lijst(20, x)
        
        For Each pp In lijstPersoneelPlanning
            If pp.personeelid = p.Id Then
            p.PersoneelPlanningenToevoegen pp
            End If
        Next pp
        LijstProjectPersoneelLijstOphalen.Add p, CStr(p.Id)
    Next x

End If


End Function

Function lijstprojectenophalen() As Collection
Dim lijst As Variant
Dim p As project
Set lijstprojectenophalen = New Collection
Dim l As Long
Dim db As New DataBase
lijst = db.getLijstBySQL("SELECT * FROM PROJECTEN")

For l = 0 To UBound(lijst, 2)
    Set p = New project
    p.FromList l, lijst
    lijstprojectenophalen.Add p, p.synergy & "-" & p.Vestiging
Next l

End Function

Function laatsterij() As Long
    laatsterij = ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_achternaam & 1048576).End(xlUp).Row
End Function

Sub VerwijderPersoneelPlanning()
Attribute VerwijderPersoneelPlanning.VB_ProcData.VB_Invoke_Func = "d\n14"
If Blad5.Name = ActiveSheet.Name And (Environ("UserName") = "r.bergevoet" Or Environ("UserName") = "c.arink" Or Environ("UserName") = "r.harmsen" Or Environ("UserName") = "r.kramp" Or Environ("UserName") = "g.vanderveen" Or Environ("UserName") = "Roderik") Then
    PersoneelsPlanning.DeletePP Selection
ElseIf Blad4.Name = ActiveSheet.Name Then
    If IsNumeric(ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & ActiveCell.Row)) = True Then
        MaterielenPlanning.DeleteMP Selection
    End If
End If

End Sub

Sub ToevoegenPersoneelsPlanning()
Attribute ToevoegenPersoneelsPlanning.VB_ProcData.VB_Invoke_Func = "i\n14"
If Blad5.Name = ActiveSheet.Name And (Environ("UserName") = "r.bergevoet" Or Environ("UserName") = "c.arink" Or Environ("UserName") = "r.harmsen" Or Environ("UserName") = "r.kramp" Or Environ("UserName") = "g.vanderveen" Or Environ("UserName") = "Roderik") Then
    FORM_PROJECT_KIEZER.Show
ElseIf Blad4.Name = ActiveSheet.Name Then
    If IsNumeric(ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & ActiveCell.Row)) = True And ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & ActiveCell.Row) <> "" Then
        FORM_PROJECT_MAT.Show
    End If
End If

End Sub
