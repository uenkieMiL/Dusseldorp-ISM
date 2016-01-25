VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_KANTOORPERSONEEL 
   Caption         =   "BEHEREN KANTOORPERSONEEL"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690.001
   OleObjectBlob   =   "FORM_KANTOORPERSONEEL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_KANTOORPERSONEEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CKantoorPersoneel As Collection
Private kp As KantoorPersoneel

Private Sub CommandButton1_Click()
    Set kp = New KantoorPersoneel
    kp.afkorting = TextBoxAfkorting
    kp.Naam = TextBoxNaam
    kp.Gebruikersnaam = TextBoxGebruikersnaam
    kp.Email = TextBoxEmail
    kp.pv = CheckBoxPV
    kp.pl = CheckBoxPL
    kp.CALC = CheckBoxCALC
    kp.wvb = CheckBoxWVB
    kp.uitv = CheckBoxUITV
    kp.NAB = CheckBoxNAB
    kp.OFFERTE = CheckBoxOFFERTE
    kp.Opmerking = TextBoxOpmerking
    kp.Inactief = CheckBoxInActief
    kp.insert
    Set CKantoorPersoneel = Lijsten.AlleKantoorpersoneel
    bijwerken
End Sub

Private Sub CommandButton2_Click()
    UpdateKP
    Set CKantoorPersoneel = Lijsten.AlleKantoorpersoneel
bijwerken
End Sub

Private Sub ListBox1_Click()
If ListBox1.ListIndex <> -1 Then
    Set kp = New KantoorPersoneel
    Set kp = CKantoorPersoneel.item(CStr(ListBox1.List(ListBox1.ListIndex, 0)))
    bijwerkenvelden
End If
End Sub
Private Function UpdateKP()
    Set kp = New KantoorPersoneel
    kp.afkorting = TextBoxAfkorting
    kp.Naam = TextBoxNaam
    kp.Gebruikersnaam = TextBoxGebruikersnaam
    kp.Email = TextBoxEmail
    kp.pv = CheckBoxPV
    kp.pl = CheckBoxPL
    kp.CALC = CheckBoxCALC
    kp.wvb = CheckBoxWVB
    kp.uitv = CheckBoxUITV
    kp.NAB = CheckBoxNAB
    kp.OFFERTE = CheckBoxOFFERTE
    kp.Opmerking = TextBoxOpmerking
    kp.Inactief = CheckBoxInActief
    If kp.update = True Then
        leegvelden
    Else
        Functies.errorhandler_MsgBox ("Er is iets mis gegaan met het bijwerken van de gegevens")
    End If
End Function
Private Function leegvelden()
    TextBoxAfkorting = ""
    TextBoxNaam = ""
    TextBoxGebruikersnaam = ""
    TextBoxEmail = ""
    CheckBoxPV = False
    CheckBoxPL = False
    CheckBoxCALC = False
    CheckBoxWVB = False
    CheckBoxUITV = False
    CheckBoxNAB = False
    CheckBoxOFFERTE = False
    TextBoxOpmerking = ""
    CheckBoxInActief = False
    
    If ListBox1.ListIndex <> -1 Then ListBox1.ListIndex = -1
End Function
Private Function bijwerkenvelden()
    TextBoxAfkorting = kp.afkorting
    TextBoxNaam = kp.Naam
    TextBoxGebruikersnaam = kp.Gebruikersnaam
    TextBoxEmail = kp.Email
    CheckBoxPV = kp.pv
    CheckBoxPL = kp.pl
    CheckBoxCALC = kp.CALC
    CheckBoxWVB = kp.wvb
    CheckBoxUITV = kp.uitv
    CheckBoxNAB = kp.NAB
    CheckBoxOFFERTE = kp.OFFERTE
    TextBoxOpmerking = kp.Opmerking
    CheckBoxInActief = kp.Inactief
End Function

Private Sub TextBoxZoeken_Change()
bijwerken
End Sub

Private Sub UserForm_Initialize()
Set CKantoorPersoneel = Lijsten.AlleKantoorpersoneel
bijwerken

End Sub

Private Function bijwerken()
Dim kp As KantoorPersoneel
Dim a As Long

    ListBox1.Clear
    
    If TextBoxZoeken = "" Then
        For Each kp In CKantoorPersoneel
            With ListBox1
                .AddItem
                .List(a, 0) = kp.afkorting
                .List(a, 1) = kp.Naam
            End With
            a = a + 1
        Next kp
    Else
        For Each kp In CKantoorPersoneel
            If InStr(1, kp.afkorting, TextBoxZoeken, vbTextCompare) <> 0 Or InStr(1, kp.Naam, TextBoxZoeken, vbTextCompare) <> 0 Then
                With ListBox1
                    .AddItem
                    .List(a, 0) = kp.afkorting
                    .List(a, 1) = kp.Naam
                End With
                a = a + 1
            End If
        Next kp
    End If
    
    
End Function
