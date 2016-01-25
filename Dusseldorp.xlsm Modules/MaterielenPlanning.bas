Attribute VB_Name = "MaterielenPlanning"
Option Explicit
    Public Const startkolom As Integer = 12
    Public Const startrij As Integer = 8
    Public Const col_mat_id As String = "A"
    Public Const col_mat_intern As String = "B"
    Public Const col_mat_Type As String = "C"
    Public Const col_mat_omschr As String = "D"
    Public Const col_mat_merk As String = "E"
    Public Const col_mat_bouwjaar As String = "F"
    Public Const col_mat_aanschafdatum As String = "G"
    Public Const col_mat_keuringsdatum As String = "H"
    Public Const col_mat_onderhoudstermijn As String = "I"
    Public Const col_mat_laatstekeuring As String = "J"
    Public Const col_mat_status As String = "K"
    Public Const col_plan_start As String = "L"
    Public Const col_plan_eind As String = "TP"
    Public Const row_datum As Long = 1
    Public Const row_jaar As Long = 2
    Public Const row_maand As Long = 3
    Public Const row_week As Long = 4
    Public Const row_dag As Long = 5
    Private kalender As New Collection
Function MaterieelPlanningVernieuwen()
    Dim d As datum
    Dim lk As Long
    Dim lr As Long
    Dim vandaag As Date
    Dim mindatum As Date
    Dim maxdatum As Date
    Dim r As Long
    Dim m As Materieel
    Dim materieellijst As New Collection
    Dim ws As Worksheet
    Dim mp As MaterieelPlanning
    Turbo_AAN
        If Blad4.Name <> ActiveSheet.Name Then Blad4.Select
        Range(MaterielenPlanning.col_mat_id & MaterielenPlanning.startrij).Select
        
        If Blad4.AutoFilterMode = True Then Blad4.AutoFilterMode = False
        vandaag = Now()
        
        mindatum = DateAdd("d", 0 - Weekday(vandaag, vbMonday) - 14, Now())
        maxdatum = DateAdd("d", (104 * 7 - 1), mindatum)
        
        
    ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_status & 1) = mindatum + 1
    Set kalender = Lijsten.KalenderStartEind(mindatum, maxdatum)
    Set materieellijst = MaterielenPlanning.MaakLijstMaterieel2

    lr = ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_omschr & 1048576).End(xlUp).Row
    lk = ThisWorkbook.Sheets(Blad4.Name).Range("XFD" & row_dag).End(xlToLeft).Column
    If lr > MaterielenPlanning.startrij Then ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & MaterielenPlanning.startrij & ":" & MaterielenPlanning.col_plan_eind & lr).Clear
    If lk >= MaterielenPlanning.startkolom Then ThisWorkbook.Sheets(Blad4.Name).Range(Range(Columns(MaterielenPlanning.startkolom), Columns(lk + 1)).Address).delete Shift:=xlToLeft
    
    ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.startrij, MaterielenPlanning.startkolom), Cells(lr, 16384)).Interior.Color = xlNone
    
    For Each d In kalender
        If d.Kolomnummer > -1 Then
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_datum, MaterielenPlanning.startkolom + d.Kolomnummer).Address) = d.datum
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_jaar, MaterielenPlanning.startkolom + d.Kolomnummer).Address) = Year((d.datum))
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_maand, MaterielenPlanning.startkolom + d.Kolomnummer).Address) = MonthName(Month(d.datum))
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_week, MaterielenPlanning.startkolom + d.Kolomnummer).Address) = DatePart("ww", d.datum, vbMonday, vbFirstFourDays)
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_dag, MaterielenPlanning.startkolom + d.Kolomnummer).Address) = Day(d.datum)
        ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_dag, MaterielenPlanning.startkolom + d.Kolomnummer).Address).HorizontalAlignment = xlCenter
        End If
    Next d
    lk = ThisWorkbook.Sheets(Blad4.Name).Range("XFD" & row_dag).End(xlToLeft).Column
   
    MaakRaster ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & 1 & ":" & Functies.KolomNaarLetter(CInt(lk)) & MaterielenPlanning.startrij + materieellijst.Count - 1)
    
    DagJaarSamenvoegenMateriaal
    DagMaandSamenvoegenMateriaal
    DagWeekSamenvoegenMateriaal
    
    Set ws = ThisWorkbook.Sheets(Blad4.Name)
    r = MaterielenPlanning.startrij
    For Each m In materieellijst
        ws.Range(MaterielenPlanning.col_mat_id & r) = m.Id
        ws.Range(MaterielenPlanning.col_mat_intern & r) = m.MaterieelCode
        ws.Range(MaterielenPlanning.col_mat_omschr & r) = m.Omschrijving
        ws.Range(MaterielenPlanning.col_mat_merk & r) = m.Merk
        ws.Range(MaterielenPlanning.col_mat_bouwjaar & r) = m.Bouwjaar
        If m.AanschafDatum <> #12:00:00 AM# Then ws.Range(MaterielenPlanning.col_mat_aanschafdatum & r) = m.AanschafDatum
        If m.KeuringsDatum <> #12:00:00 AM# Then ws.Range(MaterielenPlanning.col_mat_keuringsdatum & r) = m.KeuringsDatum
        ws.Range(MaterielenPlanning.col_mat_onderhoudstermijn & r) = m.Onderhoudstermijn
        If m.LaatsteOnderhoudsDatum <> #12:00:00 AM# Then ws.Range(MaterielenPlanning.col_mat_laatstekeuring & r) = m.LaatsteOnderhoudsDatum
        If m.Status = "" Then
            ws.Range(MaterielenPlanning.col_mat_status & r) = "In Magazijn"
        Else
            ws.Range(MaterielenPlanning.col_mat_status & r) = CStr(m.Status)
        End If
        For Each mp In m.CMaterieelPlanning
            If mp.KolomnummerStart > -1 Then
                With ws.Range(Cells(r, mp.KolomnummerStart + MaterielenPlanning.startkolom), Cells(r, mp.KolomnummerEind + MaterielenPlanning.startkolom))
                    .Interior.Color = mp.MaterieelSoort.Kleur
                    
                    If mp.MaterieelSoort.Koppelbaar = True Then
                        .Value = mp.synergy
                    Else
                        .Value = UCase(Left(mp.MaterieelSoort.Omschrijving, 6))
                    End If
                End With
            End If
        Next mp
        r = r + 1
    Next m
End Function

Function DagJaarSamenvoegenMateriaal()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = MaterielenPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad4.Name).Range("XFD" & MaterielenPlanning.row_jaar).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = MaterielenPlanning.startkolom + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_jaar, k).Address) <> ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_jaar, k - 1).Address) Then
        Set rng = ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(MaterielenPlanning.row_jaar, sk1), Cells(MaterielenPlanning.row_jaar, k - 1)).Address)
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
    
    Function DagMaandSamenvoegenMateriaal()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = MaterielenPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad4.Name).Range("XFD" & MaterielenPlanning.row_maand).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_maand, k).Address) <> ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_maand, k - 1).Address) Then
    
        Set rng = ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(MaterielenPlanning.row_maand, sk1), Cells(MaterielenPlanning.row_maand, k - 1)).Address)
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
    
            
        With ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(MaterielenPlanning.row_maand, sk1), Cells(MaterielenPlanning.row_maand, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
    
        sk1 = k
    End If
    
    Next k
    Application.DisplayAlerts = True
    End Function
    
    
    Function DagWeekSamenvoegenMateriaal()
    Dim lk As Long
    Dim sk1 As Long
    Dim k As Long
    Dim rng As Range
    
    sk1 = MaterielenPlanning.startkolom
    lk = ThisWorkbook.Sheets(Blad4.Name).Range("XFD" & MaterielenPlanning.row_week).End(xlToLeft).Column
    Application.DisplayAlerts = False
    For k = sk1 + 1 To lk + 1
        If ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_week, k).Address) <> ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_week, k - 1).Address) Then
    
        
            With ThisWorkbook.Sheets(Blad4.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
        
            With ThisWorkbook.Sheets(Blad4.Name).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
            End With
            
            Set rng = ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(MaterielenPlanning.row_week, sk1), Cells(MaterielenPlanning.row_week, k - 1)).Address)
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
            
        With ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(MaterielenPlanning.row_week, sk1), Cells(MaterielenPlanning.row_week, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
        
        sk1 = k
      End If
    Next k
      
End Function

Function MaakLijstMaterieel() As Collection

    Dim db As New DataBase
    Dim lijst As Variant
    Dim m As Materieel
    Dim r As Long
    
    Set MaakLijstMaterieel = New Collection
    lijst = db.getLijstBySQL("SELECT * FROM MATERIEEL WHERE Inplanbaar = True;")
    
    If IsEmpty(lijst) = False Then
        For r = 0 To UBound(lijst, 2)
            Set m = New Materieel
            m.FromList r, lijst
            MaakLijstMaterieel.Add m, CStr(m.Id)
        Next r
    End If
    
End Function

Function MaakLijstMaterieel2() As Collection

    Dim db As New DataBase
    Dim lijst As Variant, lijstMaterieel As Variant
    Dim m As Materieel
    Dim mp As MaterieelPlanning
    Dim r As Long, r2 As Long
    Dim startDatumLijst As Date
    
    startDatumLijst = ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_plan_start & MaterielenPlanning.row_datum)
    Set MaakLijstMaterieel2 = New Collection
    lijstMaterieel = db.getLijstBySQL("SELECT * from MATERIEEL WHERE Inplanbaar = true AND INACTIEF = False ORDER BY MaterieelCode;")
    lijst = db.getLijstBySQL("SELECT PLANNING_MATERIEEL.*, MATERIEELSOORT.* " & _
                             "FROM PLANNING_MATERIEEL LEFT JOIN MATERIEELSOORT ON MATERIEELSOORT.Id = PLANNING_MATERIEEL.MaterieelSoortId " & _
                             " WHERE PLANNING_MATERIEEL.StartDatum >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "# OR PLANNING_MATERIEEL.EindDatum >= #" & Month(startDatumLijst) & "/" & Day(startDatumLijst) & "/" & Year(startDatumLijst) & "# ORDER BY StartDatum;")
    
    
    If IsEmpty(lijstMaterieel) = False Then
        For r = 0 To UBound(lijstMaterieel, 2)
            Set m = New Materieel
            m.FromList r, lijstMaterieel
            If IsEmpty(lijst) = False Then
                For r2 = 0 To UBound(lijst, 2)
                    If lijst(1, r2) = m.Id Then
                        Set mp = New MaterieelPlanning
                        Set mp = GetMaterieelPlanningFromList(lijst, r2)
                        m.MaterieelPlanningenToevoegen mp
                    End If
                Next r2
            End If
            MaakLijstMaterieel2.Add m, CStr(m.Id)
        Next r
    End If
    
End Function

Function laatsterij() As Long
    laatsterij = ThisWorkbook.Sheets(Blad5.Name).Range(MaterielenPlanning.col_mat_omschr & 1048576).End(xlUp).Row
End Function

Function DeleteMP(Target As Range)
Dim c As Range
Dim mp As MaterieelPlanning
For Each c In Target
If Not Intersect(c, Range(MaterielenPlanning.col_plan_start & MaterielenPlanning.startrij - 1 & ":" & MaterielenPlanning.col_plan_eind & MaterielenPlanning.laatsterij)) Is Nothing Then
    If IsNumeric(ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_id & c.Row)) = True And c.Interior.Color <> xlNone Then
        Set mp = New MaterieelPlanning
        mp.datum = Cells(1, c.Column)
        mp.MaterieelId = Cells(c.Row, 1)
        mp.DeleteDatumMaterieel
        c.Interior.Color = xlNone
        c.ClearContents
    End If
End If

Next c

End Function



Function GetMaterieelPlanningFromList(lijst As Variant, r As Long) As MaterieelPlanning
Dim mp As New MaterieelPlanning
Dim d As New datum
Dim mindatum As Date
Dim maxdatum As Date
Dim nieuwezoekdatum As Date

    mp.Id = lijst(0, r)
    mp.MaterieelId = lijst(1, r)
    mp.startdatum = lijst(2, r)
    mp.einddatum = lijst(3, r)
    If kalender.Count = 0 Then
        mindatum = DateAdd("d", 0 - Weekday(Now(), vbMonday) - 14, Now())
        maxdatum = DateAdd("d", (104 * 7 - 1), mindatum)
        Set kalender = Lijsten.KalenderStartEind(mindatum, maxdatum)
    End If
    If InCollection(kalender, CStr(mp.startdatum)) = True Then
        Set d = kalender.item(CStr(mp.startdatum))
        If d.Kolomnummer = -1 Then
            Do Until d.Kolomnummer > -1
                nieuwezoekdatum = DateAdd("d", 1, mp.startdatum)
                If InCollection(kalender, CStr(nieuwezoekdatum)) = True Then
                    Set d = kalender.item(CStr(nieuwezoekdatum))
                    If d.Kolomnummer > -1 Then mp.KolomnummerEind = d.Kolomnummer
                End If
            Loop
        Else
            mp.KolomnummerEind = d.Kolomnummer
        End If
    Else
        mp.KolomnummerStart = -1
    End If
    
    If InCollection(kalender, CStr(mp.einddatum)) = True Then
        Set d = kalender.item(CStr(mp.einddatum))
        If d.Kolomnummer = -1 Then
            Do Until d.Kolomnummer > -1
                nieuwezoekdatum = DateAdd("d", -1, mp.einddatum)
                If InCollection(kalender, CStr(nieuwezoekdatum)) = True Then
                    Set d = kalender.item(CStr(nieuwezoekdatum))
                    If d.Kolomnummer > -1 Then mp.KolomnummerEind = d.Kolomnummer
                End If
            Loop
        Else
            mp.KolomnummerEind = d.Kolomnummer
        End If
    Else
        mp.KolomnummerEind = -1
    End If
    
    If mp.KolomnummerEind > -1 And mp.KolomnummerStart = -1 Then mp.KolomnummerStart = 0
    
    mp.MaterieelSoortId = lijst(4, r)
    mp.Gekoppeld = lijst(5, r)
    mp.synergy = lijst(6, r)
    mp.isGepickt = lijst(7, r)
    mp.MaterieelSoort.Id = lijst(8, r)
    mp.MaterieelSoort.Omschrijving = lijst(9, r)
    mp.MaterieelSoort.Kleur = lijst(10, r)
    mp.MaterieelSoort.Koppelbaar = lijst(11, r)
    
    
    Set GetMaterieelPlanningFromList = mp
End Function
