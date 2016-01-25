Attribute VB_Name = "Lijsten"
Option Explicit

Public Function KalenderGeheel() As Collection
Dim strSQL As String
Dim lijst As Variant
Dim k As datum
Dim db As New DataBase
Dim r As Long
Dim a As Long

Set KalenderGeheel = New Collection

lijst = db.getLijstBySQL("Select * from Kalender")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set k = New datum
        k.datum = lijst(0, r)
       
        k.feestdag = lijst(1, r)
        k.ExtraDag = lijst(2, r)
        If IsNull(lijst(3, r)) = False Then k.Omschrijving = lijst(3, r)
        k.Zichtbaar = lijst(4, r)
        If k.Zichtbaar = False Then
            a = a + 1
            k.Kolomnummer = -1
        Else
            k.Kolomnummer = r - a
        End If
      
        KalenderGeheel.Add k, CStr(k.datum)
    Next r
End If
Set db = Nothing
Set lijst = Nothing
End Function

Public Function KalenderOverallPlanning() As Collection
Dim strSQL As String
Dim lijst As Variant
Dim k As datum
Dim db As New DataBase
Dim r As Long
Dim a As Long

Set KalenderOverallPlanning = New Collection

lijst = db.getLijstBySQL("DAGENOVERZICHT2")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set k = New datum
        k.datum = lijst(0, r)
       
        k.feestdag = lijst(1, r)
        k.ExtraDag = lijst(2, r)
        If IsNull(lijst(3, r)) = False Then k.Omschrijving = lijst(3, r)
        k.Zichtbaar = lijst(4, r)
        If k.Zichtbaar = False Then
            a = a + 1
            k.Kolomnummer = -1
        Else
            k.Kolomnummer = r - a
        End If
      
        KalenderOverallPlanning.Add k, CStr(k.datum)
    Next r
End If
Set db = Nothing
Set lijst = Nothing
End Function

Function KalenderStartEind(mindatum As Date, maxdatum As Date) As Collection
Dim strSQL As String
Dim lijst As Variant
Dim d As datum
Dim db As New DataBase
Dim com As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim par1 As New ADODB.Parameter
Dim par2 As New ADODB.Parameter
Dim a As Long
Dim x As Long
Set KalenderStartEind = New Collection

db.Connect

With com
    .CommandText = "DAGENOVERZICHT4"     'Name of the stored procedure
    .CommandType = adCmdStoredProc  'Type : stored procedure
    .ActiveConnection = db.connection
End With

 'Create 2 output parameters
Set par1 = com.CreateParameter
par1.Name = "mindatum"
par1.Type = adDate
par1.Size = Len(mindatum)
par1.Direction = adParamInput
par1.Value = mindatum

Set par2 = com.CreateParameter
par2.Name = "maxdatum"
par2.Type = adDate
par2.Size = Len(maxdatum)
par2.Direction = adParamInput
par2.Value = maxdatum

com.Parameters.Append par1
com.Parameters.Append par2
rst.Open com

lijst = rst.GetRows

rst.Close
db.Disconnect

For x = 0 To UBound(lijst, 2)
    Set d = New datum
    
    d.datum = lijst(0, x)
    d.feestdag = lijst(3, x)
    If lijst(2, x) = True Then
        d.Kolomnummer = x - a
    Else
        a = a + 1
        d.Kolomnummer = -1
    End If
    
    KalenderStartEind.Add d, CStr(d.datum)
Next x


End Function

Function MaakSoortPlanningv2(soort As Byte) As Collection
Dim strSQL As String
Dim lijst As Variant
Dim lijst2 As Variant
Dim p As project
Dim pl As Planning
Dim pr As Productie
Dim t As taak
Dim r As Long
Dim producties As Collection
Set MaakSoortPlanningv2 = New Collection



strSQL = "SELECT PROJECTEN.*, PLANNINGEN.*, TAKEN.* " & _
         "FROM (PROJECTEN INNER JOIN PLANNINGEN ON (PROJECTEN.Synergy = PLANNINGEN.Synergy) AND (PROJECTEN.Vestiging = PLANNINGEN.Vestiging)) INNER JOIN TAKEN ON PLANNINGEN.Id = TAKEN.PlanningId " & _
         "WHERE(((planningen.soort) = " & soort & " AND PLANNINGEN.STATUS = False AND PROJECTEN.WACHT = False)) " & _
         "ORDER BY PROJECTEN.Vestiging, PROJECTEN.Synergy, TAKEN.Veld;"

lijst = Functies.LijstOpBasisVanQuery(strSQL)

strSQL = "SELECT PRODUCTIE.*, PRODUCTIESOORT.* " & _
         "FROM PRODUCTIESOORT INNER JOIN (PRODUCTIE INNER JOIN PLANNINGEN ON (PRODUCTIE.Synergy = PLANNINGEN.Synergy) AND (PRODUCTIE.Vestiging = PLANNINGEN.Vestiging)) ON PRODUCTIESOORT.Id = PRODUCTIE.Soort " & _
         "WHERE (((PLANNINGEN.Soort)= " & soort & ") AND ((PRODUCTIE.Gereed)=False)) " & _
         "ORDER BY PRODUCTIE.SYNERGY, PRODUCTIE.SOORT;"
lijst2 = Functies.LijstOpBasisVanQuery(strSQL)

If IsEmpty(lijst2) = False Then Set producties = MaakProductielijstV2(lijst2)

    If IsEmpty(lijst) = False Then
    
        For r = 0 To UBound(lijst, 2)
            If r = 0 Then
            'start
                Set p = geefUitLijstProjectTerug(lijst, r)
                p.PlanningVanProject = geefUitLijstPlanningTerug(lijst, r)
                p.PlanningVanProject.ToevoegenTaak geefUitLijstTaakTerug(lijst, r)
            ElseIf r = UBound(lijst, 2) Then
            'einde
                If p.synergy <> lijst(0, r) Then
                    'project is ongelijk (alles is onbekend)
                    Set p.PlanningVanProject = pl
                    p.CProducties = GeefCellectieProductiesVoorProject(producties, p.synergy, p.Vestiging)
                    MaakSoortPlanningv2.Add p, p.synergy
                    
                    Set p = geefUitLijstProjectTerug(lijst, r)
                    Set p.PlanningVanProject = geefUitLijstPlanningTerug(lijst, r)
                    Set t = geefUitLijstTaakTerug(lijst, r)
                    
                    p.PlanningVanProject.ToevoegenTaak t
                    p.CProducties = GeefCellectieProductiesVoorProject(producties, p.synergy, p.Vestiging)
                    MaakSoortPlanningv2.Add p, p.synergy & "-" & p.Vestiging
                Else
                    'planning en project is gelijk alleen taak toevoegen
                        Set t = geefUitLijstTaakTerug(lijst, r)
                        
                        p.PlanningVanProject.ToevoegenTaak t
                        p.CProducties = GeefCellectieProductiesVoorProject(producties, p.synergy, p.Vestiging)
                        MaakSoortPlanningv2.Add p, p.synergy & "-" & p.Vestiging
                End If
            Else
            'tussen
                If p.synergy <> lijst(0, r) Then
                    'project is ongelijk
                    p.CProducties = GeefCellectieProductiesVoorProject(producties, p.synergy, p.Vestiging)
                    MaakSoortPlanningv2.Add p, p.synergy & "-" & p.Vestiging
                    Set p = geefUitLijstProjectTerug(lijst, r)
                    p.PlanningVanProject = geefUitLijstPlanningTerug(lijst, r)

                    p.PlanningVanProject.ToevoegenTaak geefUitLijstTaakTerug(lijst, r)
                Else
                'project is gleijk
                    'planning en project is gelijk alleen taak toevoegen
                        p.PlanningVanProject.ToevoegenTaak geefUitLijstTaakTerug(lijst, r)
                End If
            End If
        Next r
    End If
            
    
End Function
Function geefUitLijstProjectTerug(ByRef lijst As Variant, r As Long) As project
    Set geefUitLijstProjectTerug = New project
    geefUitLijstProjectTerug.synergy = lijst(0, r)
    geefUitLijstProjectTerug.Omschrijving = lijst(1, r)
    geefUitLijstProjectTerug.Opdrachtgever = lijst(2, r)
    geefUitLijstProjectTerug.pv = lijst(3, r)
    geefUitLijstProjectTerug.pl = lijst(4, r)
    geefUitLijstProjectTerug.CALC = lijst(5, r)
    geefUitLijstProjectTerug.wvb = lijst(6, r)
    geefUitLijstProjectTerug.uitv = lijst(7, r)
    If IsNull(lijst(8, r)) = False Then geefUitLijstProjectTerug.NAB = lijst(8, r)
    geefUitLijstProjectTerug.OFFERTE = lijst(9, r)
    geefUitLijstProjectTerug.Vestiging = lijst(10, r)
    geefUitLijstProjectTerug.staatInWacht = lijst(12, r)
    If IsNull(lijst(13, r)) = False Then geefUitLijstProjectTerug.naBelDatum = lijst(13, r)
    If IsNull(lijst(14, r)) = False Then geefUitLijstProjectTerug.intern = lijst(14, r)
    If IsNull(lijst(15, r)) = False Then geefUitLijstProjectTerug.extern = lijst(15, r)
End Function
Function geefUitLijstPlanningTerug(ByRef lijst As Variant, r As Long) As Planning
    Set geefUitLijstPlanningTerug = New Planning
    geefUitLijstPlanningTerug.Id = lijst(16, r)
    geefUitLijstPlanningTerug.synergy = lijst(17, r)
    geefUitLijstPlanningTerug.Vestiging = lijst(18, r)
    geefUitLijstPlanningTerug.soort = lijst(19, r)
    geefUitLijstPlanningTerug.startdatum = lijst(20, r)
    geefUitLijstPlanningTerug.einddatum = lijst(21, r)
    geefUitLijstPlanningTerug.Status = lijst(22, r)
End Function
Function geefUitLijstTaakTerug(ByRef lijst As Variant, r As Long) As taak
    Set geefUitLijstTaakTerug = New taak
    geefUitLijstTaakTerug.Id = lijst(23, r)
    geefUitLijstTaakTerug.planningid = lijst(24, r)
    geefUitLijstTaakTerug.Omschrijving = lijst(25, r)
    geefUitLijstTaakTerug.Volgnummer = lijst(26, r)
    geefUitLijstTaakTerug.startdatum = lijst(27, r)
    geefUitLijstTaakTerug.einddatum = lijst(28, r)
    geefUitLijstTaakTerug.Aantal = lijst(29, r)
    geefUitLijstTaakTerug.Ehd = lijst(30, r)
    geefUitLijstTaakTerug.Status = lijst(31, r)
    geefUitLijstTaakTerug.veld = lijst(32, r)
    geefUitLijstTaakTerug.soort = lijst(33, r)
    geefUitLijstTaakTerug.BegrotingsRegel = lijst(34, r)
    If IsNull(lijst(35, r)) = False Then geefUitLijstTaakTerug.Opmerking = lijst(35, r)
    If IsNull(lijst(36, r)) = False Then geefUitLijstTaakTerug.Artikelnummer = lijst(36, r)
    geefUitLijstTaakTerug.Bestekpost = lijst(37, r)
End Function

Function MaakProductielijstV2(lijst As Variant) As Collection
    Dim pr As Productie
    Dim r As Long
    
    Set MaakProductielijstV2 = New Collection
    For r = 0 To UBound(lijst, 2)
        Set pr = New Productie
        pr.Id = lijst(0, r)
        pr.synergy = lijst(1, r)
        pr.Vestiging = lijst(2, r)
        pr.soort = lijst(3, r)
        pr.startdatum = lijst(4, r)
        pr.einddatum = lijst(5, r)
        pr.Gereed = lijst(7, r)
        pr.Omschrijving = lijst(9, r)
        pr.Kleur = lijst(10, r)
        
        MaakProductielijstV2.Add pr, CStr(pr.Id)
    Next r

End Function

Function MaakKantoorPersoneelLijst(skp As SoortKantoorPersoneel) As Collection
    Dim kp As New KantoorPersoneel
    Dim db As New DataBase
    Dim lijst As Variant
    Dim r As Long
    
    Set MaakKantoorPersoneelLijst = New Collection
    Select Case skp
    
    Case SoortKantoorPersoneel.pv
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where PV = True;")
    Case SoortKantoorPersoneel.pl
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where PL = True;")
    Case SoortKantoorPersoneel.CALC
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where CALC = True;")
    Case SoortKantoorPersoneel.wvb
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where WVB = True;")
    Case SoortKantoorPersoneel.uitv
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where UITV = True;")
    Case SoortKantoorPersoneel.NAB
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where NAB = True;")
    Case SoortKantoorPersoneel.OFFERTE
        lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel Where OFFERTE = True;")
    End Select
    
    If IsEmpty(lijst) = False Then
        For r = 0 To UBound(lijst, 2)
            Set kp = New KantoorPersoneel
            kp.FromList r, lijst
            MaakKantoorPersoneelLijst.Add kp, kp.afkorting
        Next r
    End If
End Function

Function AlleKantoorpersoneel() As Collection
Dim lijst As Variant
Dim db As New DataBase
Dim kp As KantoorPersoneel
Dim r As Long

Set AlleKantoorpersoneel = New Collection

lijst = db.getLijstBySQL("SELECT * FROM KantoorPersoneel WHERE INACTIEF=0 ORDER BY Afkorting")
If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set kp = New KantoorPersoneel
        kp.FromList r, lijst
        AlleKantoorpersoneel.Add kp, kp.afkorting
    Next r
End If

End Function

Function MaakLijstVestigingen() As Collection
Dim lijst As Variant
Dim v As Vestiging
Dim a As Long
Set MaakLijstVestigingen = New Collection
Dim db As New DataBase
Dim r As Long

lijst = db.getLijstBySQL("SELECT * FROM NAAM_VESTIGING;")


If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        a = a + 1
        Set v = New Vestiging
        v.Id = a
        v.Omschrijving = lijst(0, r)
        MaakLijstVestigingen.Add v, CStr(v.Id)
    Next r
End If
End Function

Function MaakLijstLocaties(synergy As String) As Collection
Dim lijst As Variant
Dim r As Long
Dim db As New DataBase
Dim l As Locatie

Set MaakLijstLocaties = New Collection

lijst = db.getLijstBySQL("Select * FROM Locaties WHERE Synergy = '" & synergy & "';")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set l = New Locatie
        l.FromList r, lijst
        MaakLijstLocaties.Add l, CStr(l.LocatieId)
    Next r
End If
End Function

Function MaakLijstMaterieelTypen() As Collection
Dim lijst As Variant
Dim db As New DataBase
Dim mt As MaterieelType
Dim r As Long

Set MaakLijstMaterieelTypen = New Collection

lijst = db.getLijstBySQL("Select * FROM MATERIEELTYPEN ORDER BY Artikelnummer")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        Set mt = New MaterieelType
        mt.FromList r, lijst
        MaakLijstMaterieelTypen.Add mt, CStr(mt.MaterieelTypeId)
    Next r
End If

End Function

Function MaakLijstInTePlannenMaterieel() As Collection
Dim lijst As Variant
Dim db As New DataBase
Dim o As MaterieelOrder
Dim mr As MaterieelOrderRegel
Dim r As Long

Set MaakLijstInTePlannenMaterieel = New Collection

lijst = db.getLijstBySQL("SELECT MATERIEELORDERS.*, MATERIEELORDERREGELS.*, MATERIEELTYPEN.* " & _
"FROM (MATERIEELORDERS INNER JOIN MATERIEELORDERREGELS ON MATERIEELORDERS.MaterieelOrderId = MATERIEELORDERREGELS.MaterieelOrderId) INNER JOIN MATERIEELTYPEN ON MATERIEELORDERREGELS.MaterieelTypeId = MATERIEELTYPEN.MaterieelTypeId " & _
"WHERE (((MATERIEELORDERS.Status) = 1)) " & _
"ORDER BY MATERIEELORDERS.MaterieelOrderId;")

If IsEmpty(lijst) = False Then
    For r = 0 To UBound(lijst, 2)
        If r = 0 Then
        'start
            Set o = New MaterieelOrder
            o.FromList r, lijst
            Set mr = New MaterieelOrderRegel
            mr.FromListExtra r, lijst, 10
            mr.MaterieelType.FromList r, lijst, 20
            o.ToevoegencOrderregels mr
        ElseIf r = UBound(lijst, 2) Then
        'laatste regel.
            'check if orderid is gelijk
            If o.MaterieelOrderId = lijst(0, r) Then
                'gelijk dus voeg alleen materieelregeltoe en voeg order aan lijst (laatste)
                    Set mr = New MaterieelOrderRegel
                    mr.FromListExtra r, lijst, 10
                    mr.MaterieelType.FromList r, lijst, 20
                    o.ToevoegencOrderregels mr
                MaakLijstInTePlannenMaterieel.Add o, CStr(o.MaterieelOrderId)
            Else
                'niet gelijk dus gelijk als start alles aanmaken en extra toevoegen aan lijst
                MaakLijstInTePlannenMaterieel.Add o, CStr(o.MaterieelOrderId)
                Set o = New MaterieelOrder
                o.FromList r, lijst
                    Set mr = New MaterieelOrderRegel
                    mr.FromListExtra r, lijst, 10
                    mr.MaterieelType.FromList r, lijst, 20
                    o.ToevoegencOrderregels mr
            End If
        Else
        ' tussen begin en eind
            If o.MaterieelOrderId = lijst(0, r) Then
                'gelijk dus voeg alleen materieelregeltoe en voeg order aan lijst (laatste)
                    Set mr = New MaterieelOrderRegel
                    mr.FromListExtra r, lijst, 10
                    mr.MaterieelType.FromList r, lijst, 20
                    o.ToevoegencOrderregels mr
            Else
                'niet gelijk dus gelijk als start alles aanmaken en extra toevoegen aan lijst
                MaakLijstInTePlannenMaterieel.Add o, CStr(o.MaterieelOrderId)
                Set o = New MaterieelOrder
                o.FromList r, lijst
                
                    Set mr = New MaterieelOrderRegel
                    mr.FromListExtra r, lijst, 10
                    mr.MaterieelType.FromList r, lijst, 20
                    o.ToevoegencOrderregels mr
            End If
        End If
    Next r
End If

End Function

