VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PERSONEEL_OVERZICHT 
   Caption         =   "Personeel Beheren"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
   OleObjectBlob   =   "FORM_PERSONEEL_OVERZICHT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PERSONEEL_OVERZICHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public b As New Bedrijf
Public lijst As Collection

Private Sub ComboBoxBeoordeling_Change()
updateBeoordeling
End Sub

Private Sub CommandButton1_Click()

Dim p As New Personeel
If ListBox1.ListIndex <> -1 Then
    p.Id = ListBox1.List(ListBox1.ListIndex, 0)
    p.Achternaam = TextBoxAchternaam
    p.Naam = TextBoxVoornaam
    p.BSN = TextBoxBSN
    p.Machinist = CheckBoxMachinist
    p.Timmerman = CheckBoxTimmerman
    p.Grondwerker = CheckBoxGrondwerker
    p.Uitvoerder = CheckBoxUitvoerder
    p.Sloper = CheckBoxSloper
    p.DHV = CheckBoxDAV
    p.DTA = CheckBoxDTA
    p.KVP = CheckBoxKVP
    p.HVK = CheckBoxHVK
    p.Beoordeling = ComboBoxBeoordeling.Value
    p.Bijzonderheden = TextBoxOpmerking
    p.Archief = CheckBoxArchief
    p.BedrijfId = TextBoxBedrijfId
    p.save
    
    ListBox1.List(ListBox1.ListIndex, 1) = p.Achternaam
    ListBox1.List(ListBox1.ListIndex, 2) = p.Naam
    ListBox1.List(ListBox1.ListIndex, 3) = p.BSN
    ListBox1.List(ListBox1.ListIndex, 4) = TextBoxBedrijf
    
End If
End Sub

Private Sub CommandButton2_Click()
    If TextBoxAchternaam = "" Or TextBoxVoornaam = "" Or ComboBoxBeoordeling = "" Or TextBoxBedrijfId = "" Then
        MsgBox "U kunt de persoon niet toevoegen door de mogelijk volgende redenen:" & vbNewLine & _
        " - er is geen voornaam opgegeven" & vbNewLine & _
        " - er is geen achternaam opgegeven" & vbNewLine & _
        " - er is geen beoordeling opgegeven" & vbNewLine & _
        " - er is geen bedrijf geselecteerd" & vbNewLine, vbCritical, "Fout bij het aan maken"
        Exit Sub
    End If
    
    Set p = New Personeel
    p.Achternaam = TextBoxAchternaam
    p.Naam = TextBoxVoornaam
    If TextBoxBSN <> "" Or IsNumeric(TextBoxBSN) = True Then p.BSN = TextBoxBSN
    p.Machinist = CheckBoxMachinist
    p.Timmerman = CheckBoxTimmerman
    p.Grondwerker = CheckBoxGrondwerker
    p.Uitvoerder = CheckBoxUitvoerder
    p.Sloper = CheckBoxSloper
    p.DHV = CheckBoxDAV
    p.DTA = CheckBoxDTA
    p.KVP = CheckBoxKVP
    p.HVK = CheckBoxHVK
    p.Beoordeling = ComboBoxBeoordeling.Value
    p.Bijzonderheden = TextBoxOpmerking
    p.Archief = CheckBoxArchief
    p.BedrijfId = TextBoxBedrijfId
    p.save
    
    leegvelden
    Set lijst = getpersoneelijst
    BijwerkenPersoneel
End Sub

Private Sub CommandButton3_Click()
leegvelden
End Sub
Function leegvelden()
TextBoxVoornaam = ""
TextBoxAchternaam = ""
TextBoxBSN = ""
TextBoxBedrijf = ""
TextBoxBedrijfId = ""
CheckBoxMachinist = False
CheckBoxTimmerman = False
CheckBoxGrondwerker = False
CheckBoxUitvoerder = False
CheckBoxSloper = False
CheckBoxDAV = False
CheckBoxDTA = False
CheckBoxKVP = False
CheckBoxHVK = False
ComboBoxBeoordeling = ""
TextBoxOpmerking = ""
CheckBoxArchief = False
updateBeoordeling
If ListBox1.ListIndex <> -1 Then ListBox1.Selected(ListBox1.ListIndex) = False
End Function

Private Sub CommandButton4_Click()
FORM_BEDRIJF_BEHEREN.Show
End Sub

Private Sub ListBox1_Click()
Dim p As Personeel
Set p = New Personeel
p.GetById (CLng(ListBox1.List(ListBox1.ListIndex, 0)))
setPersoneelData p
End Sub

Function setPersoneelData(p As Personeel)
TextBoxVoornaam = p.Naam
TextBoxAchternaam = p.Achternaam
TextBoxBSN = p.BSN
TextBoxBedrijf = p.Bedrijf.Bedrijfsnaam
TextBoxBedrijfId = p.BedrijfId
CheckBoxMachinist = p.Machinist
CheckBoxTimmerman = p.Timmerman
CheckBoxGrondwerker = p.Grondwerker
CheckBoxUitvoerder = p.Uitvoerder
CheckBoxSloper = p.Sloper
CheckBoxDAV = p.DHV
CheckBoxDTA = p.DTA
CheckBoxHVK = p.HVK
CheckBoxKVP = p.KVP
ComboBoxBeoordeling = p.Beoordeling
TextBoxOpmerking = p.Bijzonderheden
CheckBoxArchief = p.Archief
updateBeoordeling
End Function
Function updateBeoordeling()
Dim Kleur As Long
Select Case ComboBoxBeoordeling.Value

    Case ""
        Kleur = -2147483643
    Case 1
        Kleur = 5296274
    Case 2
        Kleur = 49407
    Case 3
        Kleur = -2147483643
    Case 4
        Kleur = 255
End Select

TextBoxVoornaam.BackColor = Kleur
TextBoxAchternaam.BackColor = Kleur
ComboBoxBeoordeling.BackColor = Kleur
End Function

Private Sub TextBoxZoeken_Change()
    BijwerkenPersoneel
End Sub

Private Sub UserForm_Initialize()
Dim a As Long


With ComboBoxBeoordeling
.AddItem
.List(0, 0) = 1
.List(0, 1) = "Goed"
.AddItem
.List(1, 0) = 2
.List(1, 1) = "Voldoende"
.AddItem
.List(2, 0) = 3
.List(2, 1) = "Onbekend"
.AddItem
.List(3, 0) = 4
.List(3, 1) = "Onvoldoende"
End With
Set lijst = getpersoneelijst
BijwerkenPersoneel

End Sub

Function getpersoneelijst() As Collection
Dim p As Personeel
Dim b As Bedrijf
Dim lijstp As Variant
Dim db As New DataBase


Set getpersoneelijst = New Collection

lijstp = db.getLijstBySQL("SELECT PERSONEEL.*, BEDRIJVEN.* FROM BEDRIJVEN INNER JOIN PERSONEEL ON BEDRIJVEN.Id = PERSONEEL.BedrijfId ORDER BY BEDRIJVEN.Bedrijfsnaam, PERSONEEL.Achternaam;")

For x = 0 To UBound(lijstp, 2)
    Set p = New Personeel
    p.Id = lijstp(0, x)
    p.Achternaam = lijstp(1, x)
    p.Naam = lijstp(2, x)
    p.BSN = lijstp(3, x)
    p.Machinist = lijstp(4, x)
    p.Timmerman = lijstp(5, x)
    p.Grondwerker = lijstp(6, x)
    p.Sloper = lijstp(7, x)
    p.DHV = lijstp(8, x)
    p.DTA = lijstp(9, x)
    p.Uitvoerder = lijstp(10, x)
    p.Bijzonderheden = lijstp(11, x)
    p.Beoordeling = lijstp(12, x)
    p.Archief = lijstp(13, x)
    p.BedrijfId = lijstp(14, x)
    p.KVP = lijstp(15, x)
    p.HVK = lijstp(16, x)
    p.Bedrijf.Id = lijstp(17, x)
    p.Bedrijf.KVK = lijstp(18, x)
    p.Bedrijf.Bedrijfsnaam = lijstp(19, x)
    If IsNull(lijstp(20, x)) = False Then p.Bedrijf.Contactpersoon = lijstp(20, x)
    If IsNull(lijstp(21, x)) = False Then p.Bedrijf.Telefoonnummer = lijstp(21, x)
    If IsNull(lijstp(22, x)) = False Then p.Bedrijf.Emailadres = lijstp(22, x)
    getpersoneelijst.Add p, CStr(p.Id)
    
Next x

End Function

Function BijwerkenPersoneel()
Dim p As Personeel
Dim a As Long
a = 0
If ListBox1.ListCount > 0 Then ListBox1.Clear
If TextBoxZoeken = "" Then
    For Each p In lijst
        ListBox1.AddItem
        ListBox1.List(a, 0) = p.Id
        ListBox1.List(a, 1) = p.Achternaam
        ListBox1.List(a, 2) = p.Naam
        ListBox1.List(a, 3) = p.BSN
        ListBox1.List(a, 4) = p.Bedrijf.Bedrijfsnaam
        a = a + 1
    Next p
Else
For Each p In lijst
        If InStr(1, p.Achternaam, TextBoxZoeken, TextCompare) > 0 Or InStr(1, p.Naam, TextBoxZoeken, TextCompare) > 0 Or InStr(1, p.BSN, TextBoxZoeken, TextCompare) > 0 Or InStr(1, p.Bedrijf.Bedrijfsnaam, TextBoxZoeken, TextCompare) > 0 Then
            ListBox1.AddItem
            ListBox1.List(a, 0) = p.Id
            ListBox1.List(a, 1) = p.Achternaam
            ListBox1.List(a, 2) = p.Naam
            ListBox1.List(a, 3) = p.BSN
            ListBox1.List(a, 4) = p.Bedrijf.Bedrijfsnaam
            a = a + 1
        End If
    Next p

End If

End Function
