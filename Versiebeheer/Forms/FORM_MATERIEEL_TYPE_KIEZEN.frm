VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_MATERIEEL_TYPE_KIEZEN 
   Caption         =   "MATERIEEL KIEZEN"
   ClientHeight    =   5115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   OleObjectBlob   =   "FORM_MATERIEEL_TYPE_KIEZEN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_MATERIEEL_TYPE_KIEZEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lijst As New Collection

Private Sub CommandButton1_Click()
Dim r As Long
Dim fouttekst As Variant
Dim fouten As New Collection
Dim tekst As String

Set fouten = Controle

If fouten.Count > 0 Then
    tekst = "De regel kan niet worden toegevoegd aan de order om de volgende reden(en):"
    For Each fouttekst In fouten
        tekst = tekst & vbNewLine & " - " & fouttekst
    Next fouttekst
    Functies.errorhandler_MsgBox (tekst)
    Exit Sub
End If

If Functies.IsUserFormLoaded("FORM_MATERIEEL_BESTELLEN") = True Then
r = FORM_MATERIEEL_BESTELLEN.ListBoxOrderRegels.ListCount
    With FORM_MATERIEEL_BESTELLEN.ListBoxOrderRegels
    If .Enabled = False Then .Enabled = True
    .AddItem
    .List(r, 0) = ListBoxTypen.List(ListBoxTypen.ListIndex, 1)
    .List(r, 1) = ListBoxTypen.List(ListBoxTypen.ListIndex, 2)
    .List(r, 2) = TextBoxAantal
    .List(r, 3) = TextBoxStartdatum
    .List(r, 4) = TextBoxEinddatum
    .List(r, 5) = ListBoxTypen.List(ListBoxTypen.ListIndex, 0)
    End With
End If
Unload Me
End Sub

Private Sub TextBoxFilter_Change()
Bijwerkenlijst
End Sub

Private Sub TextBoxEinddatum_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBoxEinddatum <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = TextBoxEinddatum
    End If
    
    FORM_KALENDER.Show
    
    If ThisWorkbook.inladen = True Then
        TextBoxEinddatum = ThisWorkbook.datum
    End If
End Sub



Private Sub TextBoxStartdatum_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBoxStartdatum <> "" Then
        ThisWorkbook.inladen = True
        ThisWorkbook.datum = TextBoxStartdatum
    End If
    
    FORM_KALENDER.Show
    
    If ThisWorkbook.inladen = True Then
        TextBoxStartdatum = ThisWorkbook.datum
    End If
End Sub

Private Sub UserForm_Initialize()
ophalenlijst
Bijwerkenlijst
End Sub

Function ophalenlijst()
    Set lijst = Lijsten.MaakLijstMaterieelTypen
End Function

Function Bijwerkenlijst()
Dim mt As MaterieelType
Dim a As Long

ListBoxTypen.Clear

If TextBoxFilter = "" Then
    For Each mt In lijst
        If mt.Inactief = False Then
            With ListBoxTypen
                .AddItem
                .List(a, 0) = mt.MaterieelTypeId
                .List(a, 1) = mt.Artikelnummer
                .List(a, 2) = mt.Omschrijving
            End With
            
            a = a + 1
        End If
    Next mt
Else
    For Each mt In lijst
        If InStr(1, mt.Artikelnummer, TextBoxFilter, TextCompare) > 0 Or _
        InStr(1, mt.Omschrijving, TextBoxFilter, TextCompare) > 0 Then
            With ListBoxTypen
                .AddItem
                .List(a, 0) = mt.MaterieelTypeId
                .List(a, 1) = mt.Artikelnummer
                .List(a, 2) = mt.Omschrijving
            End With
            a = a + 1
        End If
    Next mt
 
End If

End Function


Function Controle() As Collection
Set Controle = New Collection

If TextBoxAantal = "" Then Controle.Add ("Er is geen aantal opgegeven")
If TextBoxStartdatum = "" Then Controle.Add ("Er is geen startdatum opgegeven")
If TextBoxEinddatum = "" Then Controle.Add ("Er is geen einddatum opgegeven")


If IsNumeric(TextBoxAantal) = False Then Controle.Add ("de waarde bij aantal is niet numeriek")
If IsDate(TextBoxStartdatum) = False Then Controle.Add ("de waarde bij startdatum is niet een datum")
If IsDate(TextBoxEinddatum) = False Then Controle.Add ("de waarde bij einddatum is niet een datum")

If ListBoxTypen.ListIndex = -1 Then Controle.Add ("er is geen materieel gekozen")
End Function
