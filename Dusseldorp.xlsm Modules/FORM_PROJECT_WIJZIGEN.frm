VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PROJECT_WIJZIGEN 
   Caption         =   "PROJECT WIJZIGEN"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8610.001
   OleObjectBlob   =   "FORM_PROJECT_WIJZIGEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PROJECT_WIJZIGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private dagen As Long
Public CKalender As Collection
Private verplaatsen As Boolean
Private formingeladen As Boolean
Public project As New project
Public newproductie As Boolean
Public prod As New Productie
Public productie_inladen As Boolean


Private Sub Check_Asbest_gemeld_Click()
If Check_Asbest_gemeld = True Then
        Combo_AsbestGemeld_Start.Visible = True
        Combo_AsbestGemeld_Eind.Visible = True
        Combo_AsbestGemeld_Start = FormatDateTime(Now(), vbShortDate)
        Combo_AsbestGemeld_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_AsbestGemeld_Start.Visible = False
        Combo_AsbestGemeld_Eind.Visible = False
    End If
End Sub

Private Sub Check_Concept_Click()
If Check_Concept = True Then
        Combo_Concept_Start.Visible = True
        Combo_Concept_Eind.Visible = True
        Combo_Concept_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Concept_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_Concept_Start.Visible = False
        Combo_Concept_Eind.Visible = False
    End If
End Sub

Private Sub Combo_AsbestGemeld_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Combo_AsbestGemeld_Eind <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = Combo_AsbestGemeld_Eind
    End If
        FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        Combo_AsbestGemeld_Eind = ThisWorkbook.datum
    End If
End Sub

Private Sub Combo_AsbestGemeld_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim startdatum As Date
startdatum = getStartDatumProductie(TextBox1, 1)

If Combo_Concept_Start <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = Combo_AsbestGemeld_Start
End If

FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    Combo_AsbestGemeld_Start = ThisWorkbook.datum
    If startdatum = #12:00:00 AM# Then dagen = False Else dagen = DateDiff("d", startdatum, Combo_AsbestGemeld_Start)
Combo_AsbestGemeld_Start = ThisWorkbook.datum
End If
End Sub



Private Sub Combo_Concept_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
   If Combo_Concept_Eind <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = Combo_Concept_Eind
    End If
        FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        Combo_Concept_Eind = ThisWorkbook.datum
    End If
End Sub

Private Sub Combo_Concept_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date
startdatum = getStartDatumProductie(TextBox1, 1)

If Combo_Concept_Start <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = Combo_Concept_Start
End If

FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    Combo_Concept_Start = ThisWorkbook.datum
    If startdatum = #12:00:00 AM# Then dagen = False Else dagen = DateDiff("d", startdatum, Combo_Concept_Start)
Combo_Concept_Start = ThisWorkbook.datum
End If
End Sub

Private Sub Combo_Uitvoering_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Combo_Uitvoering_Eind <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = Combo_Uitvoering_Eind
    End If
    FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        Combo_Uitvoering_Eind = ThisWorkbook.datum
    End If

End Sub


Private Sub Combo_Uitvoering_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Uitvoering_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Uitvoering_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Uitvoering_Start = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Acquisitie_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date
startdatum = getStartDatum(TextBox1, 1)

If Combo_Acquisitie_Start <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = Combo_Acquisitie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    Combo_Acquisitie_Start = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Acquisitie_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Acquisitie_Eind <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Acquisitie_Eind
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Acquisitie_Eind = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Calculatie_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date
startdatum = getStartDatum(TextBox1, 2)
If Combo_Calculatie_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Calculatie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    
    Combo_Calculatie_Start = ThisWorkbook.datum
    Combo_Calculatie_Start = ThisWorkbook.datum
End If

End Sub


Private Sub Combo_Calculatie_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Calculatie_Eind <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Calculatie_Eind
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Calculatie_Eind = ThisWorkbook.datum
End If

End Sub



Private Sub Combo_Asbest_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date
startdatum = getStartDatumProductie(TextBox1, 1)

If Combo_Asbest_Start <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = Combo_Asbest_Start
End If

FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    Combo_Asbest_Start = ThisWorkbook.datum
    If startdatum = #12:00:00 AM# Then dagen = False Else dagen = DateDiff("d", startdatum, Combo_Asbest_Start)
    Combo_Asbest_Start = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Totaal_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim startdatum As Date
    startdatum = getStartDatumProductie(TextBox1, 2)
    
    If Combo_Totaal_Start <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = Combo_Totaal_Start
    End If
    FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        Combo_Totaal_Start = ThisWorkbook.datum
        If startdatum = #12:00:00 AM# Then dagen = False Else dagen = DateDiff("d", startdatum, Combo_Asbest_Start)
    End If
    Combo_Totaal_Start = ThisWorkbook.datum
End Sub


Private Sub Combo_Renovatie_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date
startdatum = getStartDatumProductie(TextBox1, 3)

If Combo_Renovatie_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Renovatie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    
    Combo_Renovatie_Start = ThisWorkbook.datum
    Combo_Renovatie_Start = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Asbest_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Asbest_Eind <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Asbest_Eind
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Asbest_Eind = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Totaal_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Totaal_Eind <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Totaal_Eind
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Totaal_Eind = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Renovatie_Eind_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Renovatie_Eind <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Renovatie_Eind
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Renovatie_Eind = ThisWorkbook.datum
End If
End Sub


Private Sub Check_Aquisitie_Change()
    If Check_Aquisitie = True Then
        Combo_Acquisitie_Start.Visible = True
        Combo_Acquisitie_Eind.Visible = True
        Combo_Acquisitie_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Acquisitie_Eind = FormatDateTime(Now(), vbShortDate)
    Else
        Combo_Acquisitie_Start.Visible = Fales
        Combo_Acquisitie_Eind.Visible = False
    End If
    
    
End Sub

Private Sub Check_Calculatie_Change()
    If Check_Calculatie = True Then
        Combo_Calculatie_Start.Visible = True
        Combo_Calculatie_Eind.Visible = True
        Combo_Calculatie_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Calculatie_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_Calculatie_Start.Visible = False
        Combo_Calculatie_Eind.Visible = False
    End If
    
    
End Sub



Private Sub Check_Renovatie_Click()
    If Check_Renovatie = True Then
        Combo_Renovatie_Start.Visible = True
        Combo_Renovatie_Eind.Visible = True
        Combo_Renovatie_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Renovatie_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_Renovatie_Start.Visible = False
        Combo_Renovatie_Eind.Visible = False
    End If
End Sub

Private Sub Check_Totaal_Click()
    If Check_Totaal = True Then
        Combo_Totaal_Start.Visible = True
        Combo_Totaal_Eind.Visible = True
        Combo_Totaal_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Totaal_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_Totaal_Start.Visible = False
        Combo_Totaal_Eind.Visible = False
    End If
End Sub


Private Sub Check_Uitvoering_Change()
    If Check_Uitvoering = True Then
        Combo_Uitvoering_Start.Visible = True
        Combo_Uitvoering_Eind.Visible = True
    Else
        Combo_Uitvoering_Start.Visible = False
        Combo_Uitvoering_Eind.Visible = False
    End If
End Sub


Private Sub Check_Asbest_Change()
    If Check_Asbest = True Then
        Combo_Asbest_Start.Visible = True
        Combo_Asbest_Eind.Visible = True
        Combo_Asbest_Start = FormatDateTime(Now(), vbShortDate)
        Combo_Asbest_Eind = FormatDateTime(Now(), vbShortDate)
         Else
        Combo_Asbest_Start.Visible = False
        Combo_Asbest_Eind.Visible = False
    End If
End Sub

Function PVComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(pv)

With ComboPV
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With

End Function
Function PLComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(pl)

With ComboPL
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With


End Function
Function CALCComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(CALC)

With ComboCALC
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With
End Function
Function WVBComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(wvb)

With ComboWVB
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With

End Function


Function UITVComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(uitv)

With ComboUITV
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With
End Function

Function OFFERTEComboInladen()
Dim CPersoneel As Collection
Dim kp As New KantoorPersoneel
Dim r As Long: r = 0
Set CPersoneel = Lijsten.MaakKantoorPersoneelLijst(OFFERTE)

With ComboOfferte
    .Clear
    For Each kp In CPersoneel
        .AddItem
        .List(r, 0) = kp.afkorting
        .List(r, 1) = kp.Naam
        r = r + 1
    Next kp
End With
End Function


Function VESTComboInladen(lijst As Variant)
With Combo_Vestiging
    .Clear
    For l = 0 To UBound(lijst, 2)
        .AddItem
        .List(l, 0) = lijst(0, l)
    Next l
End With

End Function

Private Sub Combo_Acquisitie_Start_AfterUpdate()
If Combo_Acquisitie_Eind = "" Then Combo_Acquisitie_Eind = Combo_Acquisitie_Start
End Sub

Private Sub Combo_Concept_Eind_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Concept_Start_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Asbest_Eind_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Asbest_Start_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_AsbestGemeld_Start_Change()
bijwerkenUitvoering
End Sub
Private Sub Combo_AsbestGemeld_Eind_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Calculatie_Start_AfterUpdate()
If Combo_Calculatie_Eind = "" Then Combo_Calculatie_Eind = Combo_Calculatie_Start
End Sub

Private Sub Combo_Renovatie_Eind_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Renovatie_Start_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Totaal_Eind_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Totaal_Start_Change()
bijwerkenUitvoering
End Sub

Private Sub Combo_Uitvoering_Start_AfterUpdate()
If Combo_Uitvoering_Eind = "" Then Combo_Uitvoering_Eind = Combo_Uitvoering_Start
End Sub

Private Sub Combo_Productie_Start_AfterUpdate()
If Combo_Productie_Eind = "" Then Combo_Productie_Eind = Combo_Productie_Start
End Sub

Private Sub CommandButton1_Click()
Dim strSQL As String
Dim actiefout As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim fout As Variant
Dim fouten As Collection
Dim WijzigingProject As New Collection
Dim WijzigingPlanning As New Collection
Dim WijzigingProductie As New Collection
Dim aanmaken As Boolean
Dim Id As Long
Dim p As project
Dim pl As Planning
Dim pr As Productie
Dim acq As Boolean, CAL As Boolean, wvb As Boolean, uitv As Boolean
Dim concept As Boolean, asbest_ongemeld As Boolean, asbest_gemeld As Boolean, totaal As Boolean, renovatie As Boolean
Dim rng As Range

Set fouten = FoutControleAanmaken
If fouten.Count = 0 Then aanmaken = True

Set rng = ActiveCell
If aanmaken Then
        Set p = New project
        p.synergy = Label16
        p.Vestiging = Label21
        p.haalop
        
        If p.Omschrijving <> TextBox2 Then WijzigingProject.Add 1
        If p.Opdrachtgever <> TextBox3 Then WijzigingProject.Add 2
        If p.Vestiging <> Combo_Vestiging Then WijzigingProject.Add 3
        If p.pv <> ComboPV Then WijzigingProject.Add 4
        If p.pl <> ComboPL Then WijzigingProject.Add 5
        If p.CALC <> ComboCALC Then WijzigingProject.Add 6
        If p.wvb <> ComboWVB Then WijzigingProject.Add 7
        If p.uitv <> ComboUITV Then WijzigingProject.Add 8
        If p.OFFERTE <> ComboOfferte Then WijzigingProject.Add 11
        If Label16.Caption <> TextBox1.Value Then WijzigingProject.Add 9
        If Check_Wacht <> p.staatInWacht Then WijzigingProject.Add 10
        
        If WijzigingProject.Count <> 0 Then Call BijwerkenProject(WijzigingProject, p)
        
        For Each pl In p.CPlanningen
            Select Case pl.soort
            Case 1
                acq = True
                If pl.startdatum <> Combo_Acquisitie_Start Then WijzigingPlanning.Add 11
                If pl.einddatum <> Combo_Acquisitie_Eind Then WijzigingPlanning.Add 12
                If pl.Status <> Check_Acquisitie_GEREED Then WijzigingPlanning.Add 13
            
            Case 2
                CAL = True
                If pl.startdatum <> Combo_Calculatie_Start Then WijzigingPlanning.Add 21
                If pl.einddatum <> Combo_Calculatie_Eind Then WijzigingPlanning.Add 22
                If pl.Status <> Check_Calculatie_GEREED Then WijzigingPlanning.Add 23
                
            Case 4
                uitv = True
                If pl.startdatum <> Combo_Uitvoering_Start Then WijzigingPlanning.Add 41
                If pl.einddatum <> Combo_Uitvoering_Eind Then WijzigingPlanning.Add 42
                If pl.Status <> Check_Uitvoering_GEREED Then WijzigingPlanning.Add 43
            End Select

        Next pl
            If acq <> Check_Aquisitie Then If Check_Aquisitie = False Then WijzigingPlanning.Add 14 Else WijzigingPlanning.Add 10
            If CAL <> Check_Calculatie Then If Check_Calculatie = False Then WijzigingPlanning.Add 24 Else WijzigingPlanning.Add 20
            If uitv <> Check_Uitvoering Then If Check_Uitvoering = False Then WijzigingPlanning.Add 44 Else WijzigingPlanning.Add 40
            
            
        If WijzigingPlanning.Count <> 0 Then Call BijwerkenPlanningen(WijzigingPlanning)
        
        'DetailPlanning.MaakDetailPlanning
        'rng.Select
        Unload Me
        Exit Sub
Else
    output = "Het Project kan niet worden gewijzigd om de volgende redenen:"
    For Each fout In fouten
    output = output & vbNewLine & "- " & fout
    Next fout
    MsgBox output, vbCritical, "PROJECT KAN NIET WORDEN AANGEMAAKT"
End If

End Sub

Private Sub CommandButton2_Click()
Dim pv As String, pl As String, CALC As String, uitv As String, wvb As String

pv = ComboPV
pl = ComboPL
CALC = ComboCALC
wvb = ComboWVB
uitv = ComboUITV
FORM_BEHEREN_NAMEN.Show
PVComboInladen (DataBase.GetKompleteTabelLijstmetWhere("NAAM_PV", "WHERE INACTIEF=0 ORDER BY NAAM"))
PLComboInladen (DataBase.GetKompleteTabelLijstmetWhere("NAAM_PL", "WHERE INACTIEF=0 ORDER BY NAAM"))
CALCComboInladen (DataBase.GetKompleteTabelLijstmetWhere("NAAM_CAL", "WHERE INACTIEF=0 ORDER BY NAAM"))
WVBComboInladen (DataBase.GetKompleteTabelLijstmetWhere("NAAM_WVB", "WHERE INACTIEF=0 ORDER BY NAAM"))
UITVComboInladen (DataBase.GetKompleteTabelLijstmetWhere("NAAM_UITV", "WHERE INACTIEF=0 ORDER BY NAAM"))

ComboPV = pv
ComboPL = pl
ComboCALC = CALC
ComboWVB = wvb
ComboUITV = uitv
End Sub


Function bepaalStatus() As Byte
Dim Status As Byte

If Check_Uitvoering Then
    Status = 1
Else
    If Check_Aquisitie Or Check_Calculatie Then
        If Check_Aquisitie Then
            If Check_Acquisitie_GEREED = False Then Status = 2
        Else
            If Check_Calculatie_GEREED = False Then Status = 3
    
        End If
    End If
End If

bepaalStatus = Status
End Function

Function statusnaarsoort(Status As Byte) As String
    Select Case Status
    Case 1
    statusnaarsoort = "UITV"
    Case 2
    statusnaarsoort = "ACQ"
        Case 3
            statusnaarsoort = "CALC"
    End Select
End Function

Private Sub CommandButton3_Click()
Dim pr As New Productie
    FORM_PRODUCTIE_AANMAKEN.Show
    If newproductie = True Then
        With ListBox1
            .Clear
            
            For Each pr In project.CProducties
                .AddItem
                .List(.ListCount - 1, 0) = pr.Id
                .List(.ListCount - 1, 1) = pr.Omschrijving
                .List(.ListCount - 1, 2) = pr.startdatum
                .List(.ListCount - 1, 3) = pr.einddatum
                If pr.Gereed = True Then .List(.ListCount - 1, 4) = "X" Else .List(.ListCount - 1, 4) = ""
            Next pr
        End With
        
        bijwerkenUitvoeringperiode
    End If
    
End Sub

Private Sub CommandButton4_Click()
Dim pr As New Productie
Dim pro As Productie
    If ListBox1.ListIndex <> -1 Then
        pr.Id = ListBox1.List(ListBox1.ListIndex, 0)
        If pr.delete = True Then
            i = ListBox1.ListIndex
            'ListBox1.RemoveItem i
            project.CProducties.Remove (i + 1)
            
            With ListBox1
                .Clear
                
                For Each pr In project.CProducties
                    .AddItem
                    .List(.ListCount - 1, 0) = pr.Id
                    .List(.ListCount - 1, 1) = pr.Omschrijving
                    .List(.ListCount - 1, 2) = pr.startdatum
                    .List(.ListCount - 1, 3) = pr.einddatum
                    If pr.Gereed = True Then .List(.ListCount - 1, 4) = "X" Else .List(.ListCount - 1, 4) = ""
                Next pr
            End With
            
            bijwerkenUitvoeringperiode
        End If
    End If
End Sub

Private Sub CommandButton5_Click()
    Dim pr As New Productie
    If ListBox1.ListIndex <> -1 Then
        pr.Id = ListBox1.List(ListBox1.ListIndex, 0)
        pr.haalop
        If pr.Gereed = True Then pr.Gereed = False Else pr.Gereed = True
        
            If pr.update = True Then
                If pr.Gereed = True Then ListBox1.List(ListBox1.ListIndex, 4) = "X" Else ListBox1.List(ListBox1.ListIndex, 4) = ""
            End If
        End If
        
End Sub

Private Sub CommandButton6_Click()
    Dim pr As New Productie
    productie_inladen = False
    If ListBox1.ListIndex <> -1 Then
        Set prod = New Productie
        prod.Id = ListBox1.List(ListBox1.ListIndex, 0)
        prod.haalop
        productie_inladen = True
        FORM_PRODUCTIE_AANPASSEN.Show
    End If
    
End Sub




'aanmakenTakenProductie(bepaalStatus,t.id)

Private Sub UserForm_Initialize()

Dim p As project
Dim pl As Planning
Dim pr As Productie
Dim soort As Long
Dim Id As String
Dim Vestiging As String
Dim db As New DataBase
Set CKalender = Lijsten.KalenderGeheel
Id = ThisWorkbook.synergy_id
Vestiging = ThisWorkbook.Vestiging

PVComboInladen
PLComboInladen
CALCComboInladen
WVBComboInladen
UITVComboInladen
OFFERTEComboInladen
VESTComboInladen (db.getLijstBySQL("SELECT * FROM NAAM_VESTIGING"))

Set p = New project
p.synergy = Id
p.Vestiging = Vestiging
p.haalop
If Not p Is Nothing Then Set project = p
TextBox1 = p.synergy
Label16.Caption = p.synergy
Label21.Caption = p.Vestiging
TextBox2 = p.Omschrijving
TextBox3 = p.Opdrachtgever
ComboPV = p.pv
ComboPL = p.pl
ComboCALC = p.CALC
ComboWVB = p.wvb
ComboUITV = p.uitv
ComboOfferte = p.OFFERTE
Combo_Vestiging = p.Vestiging
Check_Wacht = p.staatInWacht
If p.staatInWacht = True Then
    TextBox4.Visible = True
    TextBox4 = p.naBelDatum
End If

For Each pl In p.CPlanningen
    
    Select Case pl.soort
    Case 1
        Check_Aquisitie = True
        Combo_Acquisitie_Start = pl.startdatum
        Combo_Acquisitie_Eind = pl.einddatum
        Check_Acquisitie_GEREED.Visible = True
        If pl.Status = True Then Check_Acquisitie_GEREED = True
        
    Case 2
        Check_Calculatie = True
        Combo_Calculatie_Start = pl.startdatum
        Combo_Calculatie_Eind = pl.einddatum
        Check_Calculatie_GEREED.Visible = True
        If pl.Status = True Then Check_Calculatie_GEREED = True
        
    Case 4
        Check_Uitvoering = True
        Combo_Uitvoering_Start = pl.startdatum
        Combo_Uitvoering_Eind = pl.einddatum
        Check_Uitvoering_GEREED.Visible = True
        If pl.Status = True Then Check_Uitvoering_GEREED = True
    End Select
    
Next pl
    
For Each pr In p.CProducties
    With ListBox1
        .AddItem
        .List(.ListCount - 1, 0) = pr.Id
        .List(.ListCount - 1, 1) = pr.Omschrijving
        .List(.ListCount - 1, 2) = pr.startdatum
        .List(.ListCount - 1, 3) = pr.einddatum
        If pr.Gereed = True Then .List(.ListCount - 1, 4) = "X" Else .List(.ListCount - 1, 4) = ""
        
    End With
Next pr
formingeladen = True

End Sub


Function FoutControleAanmaken() As Collection
Dim fouten As Collection
Set fouten = New Collection

If Check_Aquisitie = True Then
    If Combo_Acquisitie_Start = "" Then
        fouten.Add "De startdatum van de acquisitie is niet ingevuld"
    Else
        If Not IsDate(Combo_Acquisitie_Start) Then fouten.Add "De startdatum van de acquisitie is geen datum"
    End If
    
    If Combo_Acquisitie_Eind = "" Then
        fouten.Add "De einddatum van de acquisitie is niet ingevuld"
    Else
        If Not IsDate(Combo_Acquisitie_Eind) Then fouten.Add "De startdatum van de acquisitie is geen datum"
    End If
End If

If Check_Calculatie = True Then
    If Combo_Calculatie_Start = "" Then
        fouten.Add "De startdatum van de calculatie is niet ingevuld"
    Else
    If Not IsDate(Combo_Calculatie_Start) Then fouten.Add "De startdatum van de calculatie is geen datum"
    End If
    
    If Combo_Calculatie_Eind = "" Then
        fouten.Add "De einddatum van de calculatie is niet ingevuld"
    Else
        If Not IsDate(Combo_Calculatie_Eind) Then fouten.Add "De einddatum van de calculatie is geen datum"
    End If
End If

If Check_Uitvoering = True Then
    If Combo_Uitvoering_Start = "" Then
        fouten.Add "De startdatum van de uitvoering is niet ingevuld"
    Else
        If Not IsDate(Combo_Uitvoering_Start) Then fouten.Add "De startdatum van de uitvoering is geen datum"
    End If
    
    If Combo_Uitvoering_Eind = "" Then
        fouten.Add "De einddatum van de uitvoering is niet ingevuld"
    Else
        If Not IsDate(Combo_Uitvoering_Eind) Then fouten.Add "De einddatum van de uitvoering is geen datum"
    End If
End If

    If Not Check_Aquisitie And Not Check_Calculatie And Not Check_Uitvoering Then fouten.Add "Er is geen planning opgegeven om aan te maken"
    
Set FoutControleAanmaken = fouten
End Function
Function CheckPlanningIsAangemaakt(synergy As String, soort As String) As Boolean
Dim lijst As Variant
lijst = DataBase.LijstOpBasisVanQuery("SELECT Count(*) AS AANTAL FROM PROJECTEN WHERE (((PROJECTEN.Synergy)='" & synergy & "'));")
 If lijst(0, 0) > 0 Then CheckPlanningIsAangemaakt = True
End Function
Function CheckProductieIsAangemaakt(synergy As String, soort As String) As Boolean
Dim lijst As Variant
lijst = DataBase.LijstOpBasisVanQuery("SELECT Count(*) AS AANTAL FROM PRODUCTIE WHERE Synergy='" & synergy & "'));")
 If lijst(0, 0) > 0 Then CheckProductieIsAangemaakt = True
End Function

Function BijwerkenProject(lijst As Collection, p As project)
Dim Code As Variant
'Dim Log As New Log
'Dim logtekst As Log
Dim bijwerken As Boolean

For Each Code In lijst
    Select Case Code
    Case 1, 2, 4, 5, 6, 7, 8, 10, 11
        bijwerken = True
    
    Case 3
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Vestiging", Combo_Vestiging, "PROJECTEN", Label21)
    Call UpdateUitvoerenViaTabelEnVeld(Label16, "Vestiging", Combo_Vestiging, "BEGROTINGREGELS")
    Call UpdateUitvoerenViaTabelEnVeld(Label16, "Vestiging", Combo_Vestiging, "BEHOEFTEN")
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Vestiging", Combo_Vestiging, "PLANNINGEN", Label21)
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Vestiging", Combo_Vestiging, "PRODUCTIE", Label21)

    Case 9
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Synergy", TextBox1, "PROJECTEN", Label21)
    Call UpdateUitvoerenViaTabelEnVeld(Label16, "Synergy", TextBox1, "BEGROTINGREGELS")
    Call UpdateUitvoerenViaTabelEnVeld(Label16, "Synergy", TextBox1, "BEHOEFTEN")
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Synergy", TextBox1, "PLANNINGEN", Label21)
    Call UpdateUitvoerenViaTabelEnVeldv2(Label16, "Synergy", TextBox1, "PRODUCTIE", Label21)

     
    End Select
    

Next Code

If bijwerken = True Then
    p.Omschrijving = TextBox2
    p.Opdrachtgever = TextBox3
    p.pv = ComboPV
    p.pl = ComboPL
    p.CALC = ComboCALC
    p.wvb = ComboWVB
    p.uitv = ComboUITV
    p.OFFERTE = ComboOfferte
    p.staatInWacht = Check_Wacht
    p.update
End If
End Function


Function BijwerkenPlanningen(lijst As Collection)
Dim Code As Variant
Dim Id As Long
Dim datum As Date
Dim planningid As Long
Dim aantaldagen As Integer
Dim aantalwerkdagen As Integer
Dim updateACQ As Boolean
Dim updateCAL As Boolean
Dim updateWVB As Boolean
Dim updateUITV As Boolean
Dim pl As New Planning

    For Each Code In lijst
        Select Case Code
        Case 10
            Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 1
            pl.startdatum = Combo_Acquisitie_Start
            pl.einddatum = Combo_Acquisitie_Eind
            pl.Create
            pl.VoegTakenToe
           
        Case 11
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 1, Combo_Vestiging)
            aantaldagen = DateDiff("d", pl.startdatum, CDate(Combo_Acquisitie_Start))
            If aantaldagen <> 0 Then
                pl.UpdateTakenNaarVerplaatsendatumPlanning aantaldagen
            End If
            pl.startdatum = Combo_Acquisitie_Start
            pl.update
            updateCAL = True
            
        Case 12
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 1, Combo_Vestiging)
            pl.einddatum = Combo_Acquisitie_Eind
            pl.update
            
        Case 13
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 1, Combo_Vestiging)
            pl.Status = Check_Acquisitie_GEREED
            pl.update
                
        Case 14
            DeletePlanningTaken getPlanningID(TextBox1, 1, Combo_Vestiging), 1
            
        Case 20
            Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 2
            pl.startdatum = Combo_Calculatie_Start
            pl.einddatum = Combo_Calculatie_Eind
            pl.Create
            pl.VoegTakenToe
        
        Case 21
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 2, Combo_Vestiging)
            aantaldagen = DateDiff("d", pl.startdatum, CDate(Combo_Calculatie_Start))
            If aantaldagen <> 0 Then
                pl.UpdateTakenNaarVerplaatsendatumPlanning aantaldagen
            End If
            pl.startdatum = Combo_Calculatie_Start
            pl.update
            updateCAL = True
            
        Case 22
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 2, Combo_Vestiging)
            pl.einddatum = Combo_Calculatie_Eind
            pl.update
            
        Case 23
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 2, Combo_Vestiging)
            pl.Status = Check_Calculatie_GEREED
            pl.update
        
        Case 24
            DeletePlanningTaken getPlanningID(TextBox1, 2, Combo_Vestiging), 2
        
        Case 40
            Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 4
            pl.startdatum = Combo_Uitvoering_Start
            pl.einddatum = Combo_Uitvoering_Eind
            pl.Create
            pl.VoegTakenToe
        
        Case 41
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 4, Combo_Vestiging)
            aantaldagen = DateDiff("d", pl.startdatum, CDate(Combo_Uitvoering_Start))
            If aantaldagen <> 0 Then
                pl.UpdateTakenNaarVerplaatsendatumPlanning aantaldagen
            End If
            pl.startdatum = Combo_Uitvoering_Start
            pl.update
            updateUITV = True
        Case 42
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 4, Combo_Vestiging)
            pl.einddatum = Combo_Uitvoering_Eind
            pl.update
            
        Case 43
            Set pl = New Planning
            pl.haalop getPlanningID(TextBox1, 4, Combo_Vestiging)
            pl.Status = Check_Uitvoering_GEREED
            pl.update
        
        Case 44
            DeletePlanningTaken getPlanningID(TextBox1, 4, Combo_Vestiging), 4
    End Select
    Next Code
    
End Function

Function UpdateUitvoerenPlanning(soort As Long, synergy As String, Vestiging As String, veld As String, Value As Variant)
Dim strSQL As String
Dim lijs As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
'Dim l As New Log
Dim logtekst As String
Dim oudewaarde As Variant
Dim Id As Long

Dim db As New DataBase


strSQL = "SELECT * FROM PLANNINGEN WHERE Soort = " & soort & "And Synergy = '" & synergy & "' And Vestiging = '" & Vestiging & "';"

db.Connect
rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
If rst.BOF = False Or rst.EOF = False Then
    oudewaarde = rst.Fields(veld)
    rst.Fields(veld) = Value
    Id = rst.Fields("Id").Value
    rst.update
    rst.MoveLast
End If
rst.Close
db.Disconnect


End Function

Function GetDataPlanning(soort As Long, synergy As String, ByRef planningid As Long, ByRef datum As Date)
Dim strSQL As String
Dim lijs As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn

strSQL = "SELECT * FROM PLANNINGEN WHERE Soort = " & soort & "And Synergy = '" & synergy & "';"

rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
planningid = rst.Fields("Id")
datum = rst.Fields("Einddatum")
rst.Close
cnn.Close

End Function

Function UpdateUitvoerenViaTabelEnVeld(synergy As String, veld As String, Value As String, Tabel As String)
Dim strSQL As String
Dim lijs As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim db As New DataBase

db.Connect

strSQL = "SELECT * FROM " & Tabel & " WHERE Synergy = '" & synergy & "';"

rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
If rst.BOF = False Or rst.EOF = False Then
    rst.Fields(veld) = Value
    rst.update
End If
rst.Close
db.Disconnect

End Function


Function UpdateUitvoerenViaTabelEnVeldv2(synergy As String, veld As String, Value As String, Tabel As String, Vestiging As String)
Dim strSQL As String
Dim lijs As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim db As New DataBase

db.Connect

strSQL = "SELECT * FROM " & Tabel & " WHERE Synergy = '" & synergy & "' AND Vestiging = '" & Vestiging & "';"

rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenDynamic, LockType:=adLockOptimistic
If rst.BOF = False Or rst.EOF = False Then
    rst.Fields(veld) = Value
    rst.update
End If
rst.Close
db.Disconnect
Set db = Nothing
End Function
Function bijwerkenUitvoering()
Dim startdatum As Date
Dim einddatum As Date
Dim aanpassenStartdatum As Boolean
Dim aanpassenEinddatum As Boolean
startdatum = #12/31/2099#
'If Check_Asbest = True Or Check_Totaal = True Or Check_Renovatie = True Then
'
'
'
'End If

If Check_Concept = True Then
    If IsDate(Combo_Concept_Start) = True And Combo_Concept_Start <> "" Then
        If Combo_Concept_Start < startdatum Then
            startdatum = Combo_Concept_Start
            aanpassenStartdatum = True
        End If
    End If
    
    If IsDate(Combo_Concept_Eind) = True And Combo_Concept_Eind <> "" Then
        If Combo_Concept_Eind > einddatum Then
            einddatum = Combo_Concept_Eind
            aanpassenEinddatum = True
        End If
    End If
End If

If Check_Asbest = True Then
    If IsDate(Combo_Asbest_Start) = True And Combo_Asbest_Start <> "" Then
        If Combo_Asbest_Start < startdatum Then
            startdatum = Combo_Asbest_Start
            aanpassenStartdatum = True
        End If
    End If
    
    If IsDate(Combo_Asbest_Eind) = True And Combo_Asbest_Eind <> "" Then
        If Combo_Asbest_Eind > einddatum Then
            einddatum = Combo_Asbest_Eind
            aanpassenEinddatum = True
        End If
    End If
End If

If Check_Asbest_gemeld = True Then
    If IsDate(Combo_AsbestGemeld_Start) = True And Combo_AsbestGemeld_Start <> "" Then
        If Combo_AsbestGemeld_Start < startdatum Then
            startdatum = Combo_AsbestGemeld_Start
            aanpassenStartdatum = True
        End If
    End If
    
    If IsDate(Combo_AsbestGemeld_Eind) = True And Combo_AsbestGemeld_Eind <> "" Then
        If Combo_AsbestGemeld_Eind > einddatum Then
            einddatum = Combo_AsbestGemeld_Eind
            aanpassenEinddatum = True
        End If
    End If
End If


If Check_Totaal = True Then
    If IsDate(Combo_Totaal_Start) = True And Combo_Totaal_Start <> "" Then
        If Combo_Totaal_Start < startdatum Then
            startdatum = Combo_Totaal_Start
            aanpassenStartdatum = True
        End If
    End If
    
    If IsDate(Combo_Totaal_Eind) = True And Combo_Totaal_Eind <> "" Then
        If Combo_Totaal_Eind > einddatum Then
            einddatum = Combo_Totaal_Eind
            aanpassenEinddatum = True
        End If
    End If
End If

If Check_Renovatie = True Then
    If IsDate(Combo_Renovatie_Start) = True And Combo_Renovatie_Start <> "" Then
        If Combo_Renovatie_Start < startdatum Then
            startdatum = Combo_Renovatie_Start
            aanpassenStartdatum = True
        End If
    End If
    
    If IsDate(Combo_Renovatie_Eind) = True And Combo_Renovatie_Eind <> "" Then
        If Combo_Renovatie_Eind > einddatum Then
            einddatum = Combo_Renovatie_Eind
            aanpassenEinddatum = True
        End If
    End If
End If

If aanpassenStartdatum Then Combo_Uitvoering_Start = startdatum

If aanpassenEinddatum Then Combo_Uitvoering_Eind = einddatum


End Function

Function doorschuiven(datum As Date) As Date
    Dim d As datum
    For Each d In CKalender
opnieuw:
    If d.datum = datum Then
        If d.feestdag = True Then datum = DateAdd("d", 1, datum)
        If d.Zichtbaar = False Then datum = DateAdd("d", 1, datum)
        If d.ExtraDag = True Then datum = DateAdd("d", 1, datum)
        If d.datum >= datum Then Exit For
        GoTo opnieuw:
    End If
    If d.datum > datum Then Exit For
    Next d
doorschuiven = datum
End Function
Function UpdateStartdatum(planningid As Long)
Dim strSQL As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn

strSQL = "SELECT * FROM TAKEN WHERE PlanningID = " & planningid & ";"

rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic

If rst.BOF = False Or rst.EOF = False Then
    rst.MoveFirst
    Do Until rst.EOF = True
        rst.Fields("Startdatum").Value = doorschuiven(rst.Fields("Startdatum").Value)
    rst.update
    rst.MoveNext
    Loop
    
    
End If
rst.Close
cnn.Close

End Function

Function UpdateEinddatum(planningid As Long)
Dim strSQL As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn

strSQL = "SELECT * FROM TAKEN WHERE PlanningID = " & planningid & ";"

rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenDynamic, LockType:=adLockOptimistic

If rst.BOF = False Or rst.EOF = False Then
    rst.MoveFirst
    Do Until rst.EOF = True
        rst.Fields("Startdatum").Value = doorschuiven(rst.Fields("Startdatum").Value)
    rst.update
    rst.MoveNext
    Loop
    
    
End If
rst.Close
cnn.Close

End Function

Function getPlanningID(synergy As String, soort As Byte, Vestiging As String) As Long
Dim lijst As Variant
Dim sql As String
Dim db As New DataBase

sql = "SELECT TOP 1 Id FROM PLANNINGEN WHERE synergy = '" & synergy & "' AND SOORT = " & soort & " AND Vestiging = '" & Vestiging & "';"
lijst = db.getLijstBySQL(sql)

getPlanningID = CLng(lijst(0, 0))
End Function

Function getStartDatum(synergy As String, soort As Byte) As Date
Dim lijst As Variant
Dim sql As String
sql = "SELECT TOP 1 Startdatum FROM PLANNINGEN WHERE Synergy = '" & synergy & "' AND SOORT = " & CStr(soort) & ";"
lijst = DataBase.LijstOpBasisVanQuery(sql)
If IsEmpty(lijst) = False Then
getStartDatum = lijst(0, 0)
End If

End Function

Function getStartDatumProductie(synergy As String, soort As Byte) As Date
Dim lijst As Variant
Dim sql As String
sql = "SELECT TOP 1 Startdatum from PRODUCTIE WHERE Synergy = '" & synergy & "' AND SOORT = " & CStr(soort) & ";"
lijst = DataBase.LijstOpBasisVanQuery(sql)
If IsEmpty(lijst) = False Then
    getStartDatumProductie = lijst(0, 0)
End If
End Function

Function DeletePlanningTaken(planningid As Long, soort As Byte)
    Dim antwoord As VbMsgBoxResult
    Dim sql As String
    Dim a1 As Long
    Dim a2 As Long
    Dim db As New DataBase
    
    antwoord = MsgBox("Weet u zeker dat u de planning van de " & SoortnaarString(soort) & " wilt verwijderen?", vbYesNo)
    
    If antwoord = vbYes Then
    sql = "DELETE FROM PLANNINGEN WHERE Id = " & planningid & ";"
    a1 = db.UpdateQueryUitvoeren(sql)
    sql = "DELETE FROM TAKEN WHERE PlanningId = " & planningid & ";"
    a2 = db.UpdateQueryUitvoeren(sql)
    
    MsgBox "De volgende zaken zijn verwijderd:" & vbNewLine & _
    a1 & " Planning(en)." & vbNewLine & _
    a2 & " Taken."
    
    End If
End Function

Public Function bijwerkenUitvoeringperiode()
Dim startdatum As Date
Dim einddatum As Date
Dim pr As Productie

    If Check_Uitvoering = True Then
        For Each pr In Me.project.CProducties
            If startdatum = #12:00:00 AM# Then
                startdatum = pr.startdatum
            Else
                If pr.startdatum < startdatum Then startdatum = pr.startdatum
            End If
            If einddatum = #12:00:00 AM# Then
                einddatum = pr.einddatum
            Else
                If pr.einddatum > einddatum Then einddatum = pr.einddatum
            End If
        Next pr
    
        Combo_Uitvoering_Start = startdatum
        Combo_Uitvoering_Eind = einddatum
    End If
    
    
End Function




