VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_BESTELLING_INZIEN 
   Caption         =   "UserForm1"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "FORM_BESTELLING_INZIEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_BESTELLING_INZIEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private order As MaterieelOrder
Private orderregel As MaterieelOrderRegel


Private Sub CommandButton1_Click()
    If ListBox1.ListIndex > -1 Then
        If Not orderregel Is Nothing Then
            If orderregel.Gepland = False Then
                orderregel.Gepland = True
                orderregel.GeplandDatum = Now()
                orderregel.GeplandGebruiker = Environ$("Username")
                orderregel.GeplandGebruiker = Environ$("computername")
                If orderregel.update = True Then
                    With ListBox1
                        If orderregel.Gepland = True Then .List(.ListIndex, 6) = "X"
                        If orderregel.Gepland = True Then .List(.ListIndex, 7) = orderregel.GeplandDatum
                        If orderregel.Gepland = True Then .List(.ListIndex, 8) = orderregel.GeplandGebruiker
                        If orderregel.Gepland = True Then .List(.ListIndex, 9) = orderregel.GeplandStation
                    End With
                End If
            Else
                orderregel.Gepland = False
                If orderregel.update = True Then
                    With ListBox1
                    If orderregel.Gepland = True Then .List(.ListIndex, 6) = ""
                    If orderregel.Gepland = True Then .List(.ListIndex, 7) = ""
                    If orderregel.Gepland = True Then .List(.ListIndex, 8) = ""
                    If orderregel.Gepland = True Then .List(.ListIndex, 9) = ""
                    End With
                End If
            End If
            If Functies.IsUserFormLoaded("FORM_BESTELLINGEN") = True Then
                FORM_BESTELLINGEN.HaalLijstOp
                FORM_BESTELLINGEN.Bijwerkenlijst
                Set orderregel = order.cOrderregels.item(ListBox1.List(ListBox1.ListIndex, 0))
            End If
        End If
    End If
End Sub

Private Sub CommandButton2_Click()

    With ListBox1
        If .ListIndex > -1 Then
            If .List(.ListIndex, 5) = "X" Then
                Functies.errorhandler_MsgBox ("De gekozen regel is reeds al gepland")
                Exit Sub
            Else
                FORM_BESTELLING_INZIEN.Hide
                FORM_BESTELLINGEN.Hide
                PlanRegel
            End If
        End If
    End With
End Sub

Private Function PlanRegel()
Dim rij As Long, k1 As Long, k2 As Long
Dim startdatum As Date, einddatum As Date
Dim antwoord As VbMsgBoxResult
    Dim rng As Range
    rij = Functies.SelectRegel
    startdatum = CDate(ListBox1.List(ListBox1.ListIndex, 3))
    einddatum = CDate(ListBox1.List(ListBox1.ListIndex, 4))
    k1 = Functies.DatumNaarKolomnummer(FORM_BESTELLINGEN.CKalender, startdatum)
    k2 = Functies.DatumNaarKolomnummer(FORM_BESTELLINGEN.CKalender, einddatum)
    Set rng = ThisWorkbook.Sheets(Blad4.Name).Range(Range(Cells(rij, k1 + MaterielenPlanning.startkolom), Cells(rij, k2 + MaterielenPlanning.startkolom)).Address)
    With rng
        .Interior.Color = 4886074
    End With
    
    antwooord = MsgBox()
    
    
End Function

Private Sub ListBox1_Click()

    If ListBox1.ListIndex > -1 Then
        If InCollection(order.cOrderregels, ListBox1.List(ListBox1.ListIndex, 0)) Then
            Set orderregel = order.cOrderregels.item(ListBox1.List(ListBox1.ListIndex, 0))
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    If Functies.IsUserFormLoaded("FORM_BESTELLINGEN") = True Then Set order = FORM_BESTELLINGEN.gekozenOrder
    If Not order Is Nothing Then SetForm
End Sub

Private Function SetForm()
    TextBoxAanvrager = order.Aanvrager
    TextBoxIndiener = order.Gebruiker
    TextBoxOpmerking = order.Opmerking
    TextBoxSynergy = order.synergy
    TextBoxStation = order.Station
    TextBoxTijdstip = order.Tijdstip
    CheckBoxVestiging = order.IsVestiging
    SetListbox
End Function

Function SetListbox()
    Dim mr As MaterieelOrderRegel
    
    For Each mr In order.cOrderregels
        With ListBox1
            .AddItem
            .List(.ListCount - 1, 0) = mr.MaterieelOrderRegelId
            .List(.ListCount - 1, 1) = mr.MaterieelType.Artikelnummer
            .List(.ListCount - 1, 2) = mr.MaterieelType.Omschrijving
            .List(.ListCount - 1, 3) = mr.startdatum
            .List(.ListCount - 1, 4) = mr.einddatum
            If mr.Gepland = True Then .List(.ListCount - 1, 5) = "X"
            If mr.Gepland = True Then .List(.ListCount - 1, 6) = mr.GeplandDatum
            If mr.Gepland = True Then .List(.ListCount - 1, 7) = mr.GeplandGebruiker
            If mr.Gepland = True Then .List(.ListCount - 1, 8) = mr.GeplandStation
        End With
    Next mr
End Function
