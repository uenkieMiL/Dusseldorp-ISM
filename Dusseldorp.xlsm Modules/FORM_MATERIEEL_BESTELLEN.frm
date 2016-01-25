VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_MATERIEEL_BESTELLEN 
   Caption         =   "UserForm1"
   ClientHeight    =   9255.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "FORM_MATERIEEL_BESTELLEN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_MATERIEEL_BESTELLEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private p As New project
Private kp As New KantoorPersoneel
Private afleverLocatie As Locatie
Private pMO As New MaterieelOrder



Private Sub CheckBox2_Click()
    ComboBoxBekendeLocaties.Enabled = CheckBox2.Value
End Sub


Private Sub CheckBox3_Click()
    If CheckBox3.Value = True Then
        TextBoxOmschrijving.Enabled = CheckBox3.Value
        TextBoxAdres.Enabled = CheckBox3.Value
        TextBoxPostcode.Enabled = CheckBox3.Value
        TextBoxPlaats.Enabled = CheckBox3.Value
    Else
        TextBoxOmschrijving.Enabled = CheckBox3.Value
        TextBoxAdres.Enabled = CheckBox3.Value
        TextBoxPostcode.Enabled = CheckBox3.Value
        TextBoxPlaats.Enabled = CheckBox3.Value
        TextBoxOmschrijving = ""
        TextBoxAdres = ""
        TextBoxPostcode = ""
        TextBoxPlaats = ""
    End If
End Sub

Private Sub ComboBoxBekendeLocaties_Change()

If IsNull(ComboBoxBekendeLocaties.Value) = True Then Exit Sub
    If ComboBoxBekendeLocaties.ListIndex > -1 Then
        Set afleverLocatie = New Locatie
        With ComboBoxBekendeLocaties
            afleverLocatie.LocatieId = .List(.ListIndex, 0)
            afleverLocatie.Omschrijivng = .List(.ListIndex, 1)
            afleverLocatie.Adres = .List(.ListIndex, 2)
            afleverLocatie.Postcode = .List(.ListIndex, 3)
            afleverLocatie.Plaats = .List(.ListIndex, 4)
        End With
    End If
    
    If Not afleverLocatie Is Nothing Then
        Application.EnableEvents = False
        With ComboBoxBekendeLocaties
            .Text = afleverLocatie.Omschrijivng & " | "
            .Text = .Text & afleverLocatie.Adres & " | "
            .Text = .Text & afleverLocatie.Postcode & " | "
            .Text = .Text & afleverLocatie.Plaats
        End With
        Application.EnableEvents = True
    End If
End Sub


Private Sub CommandButtonAanvragen_Click()
Dim fouten As Collection
Dim fouttekst As Variant
Dim tekst As String

Set fouten = Controle

If fouten.Count > 0 Then
    tekst = "De bestelling kan niet worden verwerkt om de volgende reden(en):"
    For Each fouttekst In fouten
        tekst = tekst & vbNewLine & " - " & fouttekst
    Next fouttekst
    Functies.errorhandler_MsgBox (tekst)
    Exit Sub
End If

If CheckBox3 = True Then
    If AanmakenLocatie = False Then
        Functies.errorhandler_MsgBox ("er is iets fout gegaan bij het aanmaken van het nieuwe adres")
        Exit Sub
    End If
End If

If AanmakenOrder = False Then
    Functies.errorhandler_MsgBox ("er is iets fout gegaan bij het aanmaken van de order")
    Exit Sub
End If

If AanmakenOrderRegels = False Then
    Functies.errorhandler_MsgBox ("er is iets fout gegaan bij het generen van de orderregels binnen de order.")
    Exit Sub
End If

Unload Me
MsgBox "De order is succesvol geplaatst. Dank voor uw bestelling.", vbInformation, "BESTELLING ONTVANGEN"

End Sub

Private Sub CommandButtonToevoegen_Click()
    FORM_MATERIEEL_TYPE_KIEZEN.Show
End Sub

Private Sub CommandButtonVerwijderen_Click()
    If ListBoxOrderRegels.ListIndex > -1 Then
        ListBoxOrderRegels.RemoveItem (ListBoxOrderRegels.ListIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    Turbo_AAN
        SetProject
        setBekendeLocaties
    turbo_UIT
End Sub


Function SetProject()
    Set p = New project
    p.synergy = ThisWorkbook.synergy_id
    p.Vestiging = ThisWorkbook.Vestiging
    p.haalop
    
    TextBoxSynergy = p.synergy
    TextBoxProjectOmschrijving = p.Omschrijving
    TextBoxOpdrachtgever = p.Opdrachtgever
    If p.uitv <> "" Then
        kp.afkorting = p.uitv
        kp.HaalOpMetAfkorting
        TextBoxAanvrager = kp.Naam
    End If
    
End Function

Function setBekendeLocaties()
Dim lijst As New Collection
Dim l As Locatie
Dim r As Long



    If Not p Is Nothing Then
        Set lijst = Lijsten.MaakLijstLocaties(p.synergy)
        
        If lijst.Count > 0 Then
        CheckBox2.Enabled = True
            For Each l In lijst
                With ComboBoxBekendeLocaties
                    .AddItem
                    .List(r, 0) = l.LocatieId
                    .List(r, 1) = l.Omschrijivng
                    .List(r, 2) = l.Adres
                    .List(r, 3) = l.Postcode
                    .List(r, 4) = l.Plaats
                End With
                
                r = r + 1

            Next l
        End If
        
    End If
End Function

Function Controle() As Collection
Set Controle = New Collection

If TextBoxSynergy = "" Then Controle.Add ("Er is geen synergy nummer opgegeven. Vul aub een geldige synergy nummer in.")

If CheckBox1 = False And CheckBox2 = False And CheckBox3 = False Then Controle.Add ("Er is geen afleverlocatie aangevinkt")

If CheckBox2 = True And ComboBoxBekendeLocaties.Value = "" Then Controle.Add ("U hebt geen reeds bekend adres aangegeven")

If CheckBox3 = True And TextBoxOmschrijving = "" Then Controle.Add ("Er is geen omschrijving van het nieuwe adres opgegeven")
If CheckBox3 = True And TextBoxAdres = "" Then Controle.Add ("Er is geen adres van het nieuwe adres opgegeven")
If CheckBox3 = True And TextBoxPostcode = "" Then Controle.Add ("Er is geen postcode van het nieuwe adres opgegeven")
If CheckBox3 = True And TextBoxPlaats = "" Then Controle.Add ("Er is geen plaats van het nieuwe adres opgegeven")

If ListBoxOrderRegels.ListCount = 0 Then Controle.Add ("Er is geen regel aangemaakt binnen deze order. Klik op het plusje om deze als nog toe te voegen.")

End Function

Private Function AanmakenLocatie() As Boolean
    Set afleverLocatie = New Locatie
    
    afleverLocatie.Omschrijivng = TextBoxOmschrijving
    afleverLocatie.Adres = TextBoxAdres
    afleverLocatie.Postcode = TextBoxPostcode
    afleverLocatie.Plaats = TextBoxPlaats
    afleverLocatie.synergy = p.synergy
    
    AanmakenLocatie = afleverLocatie.insert
    
    
End Function

Private Function AanmakenOrder() As Boolean
Dim mo As New MaterieelOrder

mo.synergy = p.synergy
mo.Aanvrager = TextBoxAanvrager
mo.Gebruiker = Environ$("username")
mo.Station = Environ$("computername")
mo.Tijdstip = Now()
If CheckBox1 = True Then
    mo.IsVestiging = True
    mo.LocatieId = 0
Else
    mo.LocatieId = afleverLocatie.LocatieId
End If

If TextBoxOpmerking.Value <> "" Then mo.Opmerking = TextBoxOpmerking.Value

mo.Status = 1

If mo.insert = True Then
    AanmakenOrder = True
    Set pMO = mo
End If
End Function

Private Function AanmakenOrderRegels()
Dim r As Long
Dim a As Long
Dim mor As MaterieelOrderRegel
Dim foutbijregels As Boolean

    With ListBoxOrderRegels
        For r = 0 To .ListCount - 1
            For a = 1 To .List(r, 2)
                Set mor = New MaterieelOrderRegel
                mor.startdatum = .List(a, 3)
                mor.einddatum = .List(a, 4)
                mor.MaterieelOrderId = pMO.MaterieelOrderId
                mor.MaterieelTypeId = .List(a, 5)
                
                If mor.insert = False Then foutbijregels = True
            Next a
        Next r
    End With
    
If foutbijregels = False Then AanmakenOrderRegels = True
End Function
