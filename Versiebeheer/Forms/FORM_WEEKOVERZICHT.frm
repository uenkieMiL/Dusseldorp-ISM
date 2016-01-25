VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_WEEKOVERZICHT 
   Caption         =   "WEEKOVERZICHT MAKEN"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3270
   OleObjectBlob   =   "FORM_WEEKOVERZICHT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_WEEKOVERZICHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ToggleButton2_Click()
Dim planningen As Collection: Set planningen = MaakCollectie2(CheckBox1.Value, Combo_Vestiging)
Dim f As Fase
Dim p As PlanningWeek
Dim pr As PlanningWeekProductie
Dim rij As Long: rij = 4
Dim rng As Range
Dim titel As Range
'Dim l As New Log
Turbo_AAN
WeekPlanning.MaakWekenlijst CheckBox1.Value, Combo_Vestiging
            
Set titel = Range("A2:J3")
    
    For Each f In planningen
    If rij <> 4 Then
        titel.Copy Range("A" & rij - 2)
        Range("B" & rij - 2) = f.Omschrijving
    End If
        For Each p In f.CProjecten
        ThisWorkbook.Sheets(Blad7.Name).Range("A" & rij) = p.synergy
        ThisWorkbook.Sheets(Blad7.Name).Range("B" & rij) = p.Omschrijving
        ThisWorkbook.Sheets(Blad7.Name).Range("C" & rij) = p.Opdrachtgever
        ThisWorkbook.Sheets(Blad7.Name).Range("D" & rij) = p.pv
        ThisWorkbook.Sheets(Blad7.Name).Range("E" & rij) = p.pl
        ThisWorkbook.Sheets(Blad7.Name).Range("F" & rij) = p.CAL
        ThisWorkbook.Sheets(Blad7.Name).Range("G" & rij) = p.wvb
        ThisWorkbook.Sheets(Blad7.Name).Range("H" & rij) = p.uitv
        
        
            Set rng = ThisWorkbook.Sheets(Blad7.Name).Range("J" & rij)
            If rng.Value = "" Then rng.Value = p.soort Else rng.Value = rng.Value & ", " & p.soort
            ThisWorkbook.Sheets(Blad7.Name).Range("I" & rij) = p.Vestiging
            
            For Each pr In p.CProductie
            
            Set rng = WeekPlanning.RijenKolommenNaarRange(rij, pr.KolomStart, pr.KolomEind)
            If Not rng Is Nothing Then
            With rng
             .Interior.Color = pr.Kleur
            End With
            End If
            Next pr
        rij = rij + 1
            
        Next p
            
        rij = rij + 3
    Next f
    'l.createLog "Weekoverzicht gegenereerd voor " & CStr(Combo_Vestiging), overzicht_gemaakt, "WEEK OVERZICHT", Overzicht
    Unload Me
    turbo_UIT
End Sub

Private Sub UserForm_Initialize()
Dim db As New DataBase
VESTComboInladen (db.getLijstBySQL("SELECT * FROM NAAM_VESTIGING"))
OptionButton1 = True
End Sub

Function MaakCollectie() As Collection
Dim projecten As Variant
Dim weeklijst As Variant
Dim productielijst As Variant
Dim planninglijst As Variant
Dim planningen As Collection
Dim plan As PlanningWeek
Dim vestigingInladen As Boolean
Dim p As Long
weeklijst = DataBase.LijstOpBasisVanQuery("WEEKOVERZICHT2")
If Combo_Vestiging <> "" Then vestigingInladen = True
If vestigingInladen = True Then
    projecten = DataBase.GetKompleteTabelLijstmetWhere("PROJECTEN", "WHERE Vestiging = '" & Combo_Vestiging & "'")
Else
    projecten = DataBase.GetKompleteTabelLijst("PROJECTEN")
End If
planninglijst = DataBase.GetKompleteTabelLijst("PLANNINGEN")
productielijst = DataBase.GetKompleteTabelLijst("PRODUCTIE")
Set planningen = New Collection
For x = 0 To UBound(projecten, 2)

        
        If CheckBox1 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox1.BackColor
            plan.soort = "ACQ"
            For p = 0 To UBound(planninglijst, 2)
                If planninglijst(2, p) = 1 And planninglijst(5, p) = False And planninglijst(1, p) = projecten(0, x) Then
                    plan.startdatum = planninglijst(3, p)
                    plan.einddatum = planninglijst(4, p)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next p
        End If
        
        If CheckBox2 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox2.BackColor
            plan.soort = "CALC"
            For p = 0 To UBound(planninglijst, 2)
                If planninglijst(2, p) = 2 And planninglijst(5, p) = False And planninglijst(1, p) = projecten(0, x) Then
                    plan.startdatum = planninglijst(3, p)
                    plan.einddatum = planninglijst(4, p)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next p
        End If
        
        If CheckBox3 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox3.BackColor
            plan.soort = "WVB"
            For p = 0 To UBound(planninglijst, 2)
                If planninglijst(2, p) = 3 And planninglijst(5, p) = False And planninglijst(1, p) = projecten(0, x) Then
                    plan.startdatum = planninglijst(3, p)
                    plan.einddatum = planninglijst(4, p)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next p
        End If
        
        If CheckBox4 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox3.BackColor
            plan.soort = "UITV"
            For p = 0 To UBound(planninglijst, 2)
                If planninglijst(3, p) = 4 And planninglijst(5, p) = False And planninglijst(1, p) = projecten(0, x) Then
                    plan.startdatum = planninglijst(3, p)
                    plan.einddatum = planninglijst(4, p)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next p
        End If
        
        If CheckBox5 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox5.BackColor
            plan.soort = "ASB"
            For prod = 0 To UBound(productielijst, 2)
                If productielijst(2, prod) = 1 And productielijst(1, prod) = projecten(0, x) Then
                    plan.startdatum = productielijst(3, prod)
                    plan.einddatum = productielijst(4, prod)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next prod
        End If
        
        If CheckBox6 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox6.BackColor
            plan.soort = "TOT"
            For prod = 0 To UBound(productielijst, 2)
                If productielijst(2, prod) = 2 And productielijst(1, prod) = projecten(0, x) Then
                    plan.startdatum = productielijst(3, prod)
                    plan.einddatum = productielijst(4, prod)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next prod
        End If
        
        If CheckBox7 = True Then
            Set plan = New PlanningWeek
            plan.synergy = projecten(0, x)
            plan.Omschrijving = projecten(1, x)
            plan.Opdrachtgever = projecten(2, x)
            plan.pv = projecten(3, x)
            plan.pl = projecten(4, x)
            plan.CAL = projecten(5, x)
            plan.wvb = projecten(6, x)
            plan.uitv = projecten(7, x)
            plan.Vestiging = projecten(8, x)
            plan.Kleur = CheckBox7.BackColor
            plan.soort = "REN"
            For prod = 0 To UBound(productielijst, 2)
                If productielijst(2, prod) = 3 And productielijst(1, prod) = projecten(0, x) Then
                    plan.startdatum = productielijst(3, prod)
                    plan.einddatum = productielijst(4, prod)
                    plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.startdatum, weeklijst)
                    plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.einddatum, weeklijst)
                    planningen.Add plan
                    Exit For
                End If
            Next prod
        End If
        
'    If CheckBox1 = True And projecten(8, x) = True And projecten(11, x) = False Then
'        Set plan = New PlanningWeek
'        plan.synergy = projecten(0, x)
'        plan.Omschrijving = projecten(1, x)
'        plan.Opdrachtgever = projecten(2, x)
'        plan.PV = projecten(3, x)
'        plan.PL = projecten(4, x)
'        plan.CAL = projecten(5, x)
'        plan.WVB = projecten(6, x)
'        plan.UITV = projecten(7, x)
'        plan.Kleur = CheckBox1.BackColor
'        plan.Soort = "ACQ"
'        plan.StartDatum = projecten(9, x)
'        plan.EindDatum = projecten(10, x)
'        plan.KolomStart = WeekPlanning.DatumNaarKolomWeek(plan.StartDatum, weeklijst)
'        plan.KolomEind = WeekPlanning.DatumNaarKolomWeek(plan.EindDatum, weeklijst)
'        planningen.Add plan
'    End If
    
   
Next x

Set MaakCollectie = planningen

End Function

Function MaakCollectie2(wacht As Boolean, Vestiging As String) As Collection
Dim projecten As Variant
Dim weeklijst As Variant
Dim productielijst As Variant
Dim planninglijst As Variant
Dim planningen As Collection
Dim plan As PlanningWeek
Dim vestigingInladen As Boolean
Dim pr As PlanningWeekProductie
Dim sql As String
Dim faselijst As Variant
Dim Fase As Fase
Dim projectaangemaakt As Boolean
Dim nf As Byte
Dim db As New DataBase

weeklijst = WeekOverzichtOphalen(wacht, Vestiging)

If OptionButton1 = True Then
    If wacht = True Then
        sql = "SELECT * FROM PROJECTEN WHERE STATUS = 0" & Vestiging
    Else
        sql = "SELECT * FROM PROJECTEN WHERE STATUS = 0 AND WACHT = 0" & Vestiging
    End If
    
    sql = sql & " ORDER BY SYNERGY"
End If

projecten = db.getLijstBySQL(sql)


productielijst = db.getLijstBySQL("SELECT * FROM PRODUCTIE ORDER BY SYNERGY, SOORT")

faselijst = db.getLijstBySQL("PRODUCTIEFASE")

Set planningen = New Collection

For f = 1 To 3
    Set Fase = New Fase
    Fase.Id = f
    Fase.Omschrijving = fasenaarstring(CByte(f))
    'ga door de faselijst en zoek projecten gelijk aan de fase. in dien gelijk voeg toe
    For fl = 0 To UBound(faselijst, 2)
        'kijk of fase gelijk is.
        If IsNull(faselijst(8, fl)) = False Then
            If IsNull(faselijst(7, fl)) = True Then faselijst(7, fl) = 0
            nf = faseomzetten(CByte(faselijst(7, fl)))
            If nf = f Then
                'ga door de projectlijst heen
                For x = 0 To UBound(projecten, 2)
                    'kijk of het synergy nummer gelijk
                    If projecten(0, x) = faselijst(0, fl) Then
                        'kijk of er al een project is aangemaakt aan de fase, zo ja zoek of het project al in de lijst staat.
                        If Fase.CProjecten.Count = 0 Then
                        'geen project dus ik mag hem toevoegen
                            Set plan = New PlanningWeek
                            plan.synergy = projecten(0, x)
                            plan.Omschrijving = projecten(1, x)
                            plan.Opdrachtgever = projecten(2, x)
                            plan.pv = projecten(3, x)
                            plan.pl = projecten(4, x)
                            plan.CAL = projecten(5, x)
                            plan.wvb = projecten(6, x)
                            plan.uitv = projecten(7, x)
                            plan.Vestiging = projecten(10, x)
                            For prod = 0 To UBound(productielijst, 2)
                                If productielijst(1, prod) = projecten(0, x) Then
                                    Set pr = New PlanningWeekProductie
                                    pr.startdatum = productielijst(4, prod)
                                    pr.einddatum = productielijst(5, prod)
                                    pr.KolomStart = WeekPlanning.DatumNaarKolomWeek(pr.startdatum, weeklijst)
                                    pr.KolomEind = WeekPlanning.DatumNaarKolomWeek(pr.einddatum, weeklijst)
                                    pr.Kleur = productielijst(5, prod)
                                    pr.soort = productielijst(2, prod)
                                    If plan.soort = "" Then
                                        plan.soort = Me.SoortnaarString(CLng(productielijst(3, prod)))
                                    Else
                                        plan.soort = plan.soort & ", " & Me.SoortnaarString(CLng(productielijst(3, prod)))
                                    End If
                                    plan.ToevoegenProductie pr
                                    
                               End If
                            
                            Next prod
                            Fase.ToevogenPlanningsweek plan
                        Else
                        'check if project al is aangemaakt in de lijst zo niet voeg hem toe.
                            projectaangemaakt = ProjectAangemaaktInCollection(Fase.CProjecten, CStr(faselijst(0, fl)))
                            If projectaangemaakt = False Then
                                Set plan = New PlanningWeek
                                plan.synergy = projecten(0, x)
                                plan.Omschrijving = projecten(1, x)
                                plan.Opdrachtgever = projecten(2, x)
                                plan.pv = projecten(3, x)
                                plan.pl = projecten(4, x)
                                plan.CAL = projecten(5, x)
                                plan.wvb = projecten(6, x)
                                plan.uitv = projecten(7, x)
                                plan.Vestiging = projecten(10, x)
                                For prod = 0 To UBound(productielijst, 2)
                                    If productielijst(1, prod) = projecten(0, x) Then
                                        Set pr = New PlanningWeekProductie
                                        pr.startdatum = productielijst(4, prod)
                                        pr.einddatum = productielijst(5, prod)
                                        pr.KolomStart = WeekPlanning.DatumNaarKolomWeek(pr.startdatum, weeklijst)
                                        pr.KolomEind = WeekPlanning.DatumNaarKolomWeek(pr.einddatum, weeklijst)
                                        pr.Kleur = productielijst(6, prod)
                                        pr.soort = productielijst(2, prod)
                                        If plan.soort = "" Then
                                            plan.soort = Me.SoortnaarString(CLng(productielijst(3, prod)))
                                        Else
                                            plan.soort = plan.soort & ", " & Me.SoortnaarString(CLng(productielijst(3, prod)))
                                        End If
                                        plan.ToevoegenProductie pr
                                        
                                   End If
                                
                                Next prod
                                Fase.ToevogenPlanningsweek plan
                            
                            End If
                        End If
                    End If
                Next x
            End If
        End If
    Next fl
    
    planningen.Add Fase
Next f

           
   
    

Set MaakCollectie2 = planningen

End Function

Function ProjectAangemaaktInCollection(projecten As Collection, synergy As String) As Boolean
Dim pw As PlanningWeek

For Each pw In projecten
    If pw.synergy = synergy Then
        ProjectAangemaaktInCollection = True
        Exit For
    End If
Next pw

End Function
Function VESTComboInladen(lijst As Variant)
With Combo_Vestiging
    .Clear
    For l = 0 To UBound(lijst, 2)
        .AddItem
        .List(l) = lijst(0, l)
    Next l
End With

End Function

Function SoortnaarString(soort As Long) As String
Select Case soort

Case 1
    SoortnaarString = "CON"
Case 2
    SoortnaarString = "REN"
Case 3
    SoortnaarString = "TOT"
Case 4
    SoortnaarString = "A-ONG"
Case 5
    SoortnaarString = "A-GEM"
Case 6
    SoortnaarString = "BOD"
End Select


End Function

Function faseomzetten(f As Byte) As Byte
Select Case f
    Case 1
     faseomzetten = 3
    Case 2
     faseomzetten = 1
    Case 3
     faseomzetten = 2
End Select
End Function

Function fasenaarstring(f As Byte) As String
Select Case f
    Case 1
     fasenaarstring = "ACQUISITIE"
    Case 2
     fasenaarstring = "CALCULATIE"
    Case 3
     fasenaarstring = "PRODUCTIE"
End Select
End Function

