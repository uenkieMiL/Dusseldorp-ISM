VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PROJECT_AANMAKEN 
   Caption         =   "PROJECT AANMAKEN"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8265.001
   OleObjectBlob   =   "FORM_PROJECT_AANMAKEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PROJECT_AANMAKEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public producties As New Collection
Public newproductie As Boolean
Public productie_inladen As Boolean


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

Private Sub Check_Uitvoering_Change()
    If Check_Uitvoering = True Then
        Combo_Uitvoering_Start.Visible = True
        Combo_Uitvoering_Eind.Visible = True
        bijwerkenUitvoeringperiode
    Else
        Combo_Uitvoering_Start.Visible = False
        Combo_Uitvoering_Eind.Visible = False
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

Private Sub Combo_Acquisitie_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Acquisitie_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Acquisitie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
    Combo_Acquisitie_Start = ThisWorkbook.datum
End If
End Sub

Private Sub Combo_Calculatie_Start_AfterUpdate()
If Combo_Calculatie_Eind = "" Then Combo_Calculatie_Eind = Combo_Calculatie_Start
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
Dim fouten As Collection
Dim aanmaken As Boolean
Dim Id As Long
Dim fout As Variant
Dim Status As Byte
Set fouten = FoutControleAanmaken
Dim statusid As Long
'Dim lp As New Log
Dim p As project
Dim pr As Productie
Dim pl As Planning
Status = bepaalStatus
Dim db As New DataBase


If fouten.Count = 0 Then aanmaken = True

If aanmaken Then
        
    db.Connect
    
    strSQL = "SELECT * FROM PROJECTEN WHERE 1=0"
    
    rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
    
        If ComboPV = "" Then ComboPV = "ONB"
        If ComboPL = "" Then ComboPL = "ONB"
        If ComboCALC = "" Then ComboCALC = "ONB"
        If ComboWVB = "" Then ComboWVB = "ONB"
        If ComboUITV = "" Then ComboUITV = "ONB"
        If ComboOfferte = "" Then ComboOfferte = "ONB"
        With rst
            .AddNew
            .Fields("Synergy").Value = TextBox1
            .Fields("Omschrijving").Value = TextBox2
            .Fields("Opdrachtgever").Value = TextBox3
            .Fields("Vestiging").Value = Combo_Vestiging
            .Fields("PV").Value = ComboPV
            .Fields("PL").Value = ComboPL
            .Fields("CALC").Value = ComboCALC
            .Fields("WVB").Value = ComboWVB
            .Fields("UITV").Value = ComboUITV
            .Fields("OFFERTE").Value = ComboOfferte
            .update
            .Close
        End With
     db.Disconnect
    
    
    
    If Check_Aquisitie = True Then
        Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 1
            pl.startdatum = Combo_Acquisitie_Start
            pl.einddatum = Combo_Acquisitie_Eind
            pl.Create
        If Status = 2 Then statusid = pl.Id
        pl.VoegTakenToe
    End If
    
    If Check_Calculatie = True Then
       Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 2
            pl.startdatum = Combo_Calculatie_Start
            pl.einddatum = Combo_Calculatie_Eind
            pl.Create
        If Status = 3 Then statusid = pl.Id
        pl.VoegTakenToe
    End If
    
    If Check_Uitvoering = True Then
        Set pl = New Planning
            pl.synergy = TextBox1
            pl.Vestiging = Combo_Vestiging
            pl.soort = 4
            pl.startdatum = Combo_Uitvoering_Start
            pl.einddatum = Combo_Uitvoering_Eind
            pl.Create
        If Status = 1 Then statusid = pl.Id
        pl.VoegTakenToe
    End If
    
    Set p = New project
    p.synergy = TextBox1
    p.Vestiging = Combo_Vestiging
    p.haalop
    'lp.createLog p.ToString, pr_aanmaken, p.synergy, project
    
    For Each pr In Me.producties
        pr.synergy = TextBox1
        pr.Vestiging = Combo_Vestiging
        pr.insert
    Next pr
    Unload Me
    Exit Sub
Else
    output = "Het Project kan niet worden aangemaakt om de volgende redenen:"
    For Each fout In fouten
    output = output & vbNewLine & "- " & fout
    Next fout
    MsgBox output, vbCritical, "PROJECT KAN NIET WORDEN AANGEMAAKT"
End If

End Sub

Private Sub CommandButton2_Click()
Dim pv As String, pl As String, CALC As String, uitv As String, wvb As String
Dim db As New DataBase
pv = ComboPV
pl = ComboPL
CALC = ComboCALC
wvb = ComboWVB
uitv = ComboUITV
FORM_KANTOORPERSONEEL.Show
PVComboInladen
PLComboInladen
CALCComboInladen
WVBComboInladen
UITVComboInladen

ComboPV = pv
ComboPL = pl
ComboCALC = CALC
ComboWVB = wvb
ComboUITV = uitv
End Sub



Private Sub CommandButton3_Click()
Dim pr As New Productie
    FORM_PRODUCTIE_AANMAKEN.Show
    If newproductie = True Then
        With ListBox1
            .Clear
            
            For Each pr In producties
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
    If ListBox1.ListIndex <> -1 Then
            i = ListBox1.ListIndex
            Me.producties.Remove (i + 1)
            
            With ListBox1
                .Clear
                
                For Each pr In producties
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

Private Sub CommandButton6_Click()
Dim prod As New Productie
    If ListBox1.ListIndex <> -1 Then
        
        Set pr = producties.item(ListBox1.ListIndex + 1)
        productie_inladen = True
        FORM_PRODUCTIE_AANPASSEN.Show
        
    End If
End Sub

Private Sub UserForm_Initialize()
Dim db As New DataBase

PVComboInladen
PLComboInladen
CALCComboInladen
WVBComboInladen
UITVComboInladen
OFFERTEComboInladen
VESTComboInladen (db.getLijstBySQL("SELECT * FROM NAAM_VESTIGING"))

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
If Combo_Calculatie_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Calculatie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
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
If Combo_Asbest_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Asbest_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Asbest_Start = ThisWorkbook.datum
End If
End Sub


Private Sub Combo_Totaal_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Totaal_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Totaal_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Totaal_Start = ThisWorkbook.datum
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


Private Sub Combo_Renovatie_Start_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Combo_Renovatie_Start <> "" Then
ThisWorkbook.inladen = True
ThisWorkbook.datum = Combo_Renovatie_Start
End If
FORM_KALENDER.Show
If ThisWorkbook.inladen = True Then
Combo_Renovatie_Start = ThisWorkbook.datum
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



Function FoutControleAanmaken() As Collection
Dim fouten As Collection
Set fouten = New Collection

If CheckProjectIsAangemaakt(TextBox1, Combo_Vestiging) = True Then fouten.Add "Project is reeds aangemaakt. Pas een bestaand planning aan."

If TextBox1 = "" Or TextBox2 = "" Or TextBox3 = "" Then
    If TextBox1 = "" Then fouten.Add "Het Synergy nummer is niet opgegeven"
    If TextBox2 = "" Then fouten.Add "De projectomschrijving is niet opgegeven"
    If TextBox3 = "" Then fouten.Add "De opdrachtgever is niet opgegeven"
End If

If Combo_Vestiging = "" Then fouten.Add "Er is geen vestiging geselecteerd"

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
    
    If ListBox1.ListCount = 0 Then fouten.Add "Er is geen productie opgegeven om aan te maken"
    
Set FoutControleAanmaken = fouten
End Function

Function bijwerkenUitvoering()
Dim startdatum As Date
Dim einddatum As Date
Dim aanpassenStartdatum As Boolean
Dim aanpassenEinddatum As Boolean
startdatum = #12/31/2099#

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

Public Function bijwerkenUitvoeringperiode()
Dim startdatum As Date
Dim einddatum As Date
Dim pr As Productie

    If Check_Uitvoering = True Then
        For Each pr In Me.producties
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
    
        If startdatum <> #12:00:00 AM# Then Combo_Uitvoering_Start = startdatum Else Combo_Uitvoering_Start = ""
        If einddatum <> #12:00:00 AM# Then Combo_Uitvoering_Eind = einddatum Else Combo_Uitvoering_Eind = ""
    End If
    
    
End Function
