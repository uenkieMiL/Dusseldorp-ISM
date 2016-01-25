VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PRODUCTIE_AANPASSEN 
   Caption         =   "PRODUCTIE WIJZIGEN"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "FORM_PRODUCTIE_AANPASSEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PRODUCTIE_AANPASSEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private pr As Productie
Private bijwerken As Boolean

Private Sub CommandButton1_Click() 'aanpassen productie
Dim Controle As New Collection
Dim fout As Variant

    Set Controle = controleer
    
    If Controle.Count = 0 And Not pr Is Nothing Then
        If Functies.IsUserFormLoaded("FORM_PROJECT_WIJZIGEN") = True Then
            pr.startdatum = CDate(TextBox1)
            pr.einddatum = CDate(TextBox2)
            pr.soort = ListBox1.List(ListBox1.ListIndex, 0)
            pr.Omschrijving = ListBox1.List(ListBox1.ListIndex, 1)
            pr.Kleur = ListBox1.List(ListBox1.ListIndex, 2)
            pr.synergy = FORM_PROJECT_WIJZIGEN.TextBox1
            pr.Vestiging = FORM_PROJECT_WIJZIGEN.Combo_Vestiging
            If pr.update = True Then
                With FORM_PROJECT_WIJZIGEN.ListBox1
                    .List(.ListIndex, 1) = pr.Omschrijving
                    .List(.ListIndex, 2) = pr.startdatum
                    .List(.ListIndex, 3) = pr.einddatum
                    If pr.Gereed = True Then .List(.ListIndex, 4) = "X" Else .List(.ListIndex, 4) = ""
                End With
                
                FORM_PROJECT_WIJZIGEN.project.CProducties.Remove (FORM_PROJECT_WIJZIGEN.ListBox1.ListIndex + 1)
                FORM_PROJECT_WIJZIGEN.project.ToevoegenProductie pr
                FORM_PROJECT_WIJZIGEN.bijwerkenUitvoeringperiode
            Else
                MsgBox "Er is iets misgegaan met het aanpassen van de productie. Probeer opnieuw.", vbCritical, "FOUT BIJ AANPASSEN PRODUCTIE."
            End If
            Me.Hide
            
        ElseIf Functies.IsUserFormLoaded("FORM_PROJECT_NIEUW") = True Then
            Set pr = FORM_PROJECT_NIEUW.producties.item(FORM_PROJECT_NIEUW.ListBox1.ListIndex + 1)
             pr.startdatum = CDate(TextBox1)
             pr.einddatum = CDate(TextBox2)
             pr.soort = ListBox1.List(ListBox1.ListIndex, 0)
             pr.Omschrijving = ListBox1.List(ListBox1.ListIndex, 1)
             pr.Kleur = ListBox1.List(ListBox1.ListIndex, 2)
             
             FORM_PROJECT_NIEUW.producties.Remove (FORM_PROJECT_NIEUW.ListBox1.ListIndex + 1)
             FORM_PROJECT_NIEUW.producties.Add item:=pr
            
             For Each pr In FORM_PROJECT_NIEUW.producties
             FORM_PROJECT_NIEUW.ListBox1.Clear
                With FORM_PROJECT_NIEUW.ListBox1
                .AddItem
                    .List(.ListCount - 1, 1) = pr.Omschrijving
                    .List(.ListCount - 1, 2) = pr.startdatum
                    .List(.ListCount - 1, 3) = pr.einddatum
                    If pr.Gereed = True Then .List(.ListCount - 1, 4) = "X" Else .List(.ListCount - 1, 4) = ""
                End With
             Next pr
             FORM_PROJECT_NIEUW.bijwerkenUitvoeringperiode
             Me.Hide
        End If
        
        
    Else
    ' productie is niet correct ingevuld
        tekst = "De Productie kan niet worden toegevoegd aan het project om de volgende reden(en):"
        For Each fout In Controle
            tekst = tekst & vbNewLine & fout
        Next fout
        MsgBox tekst, vbCritical, "FOUT BIJ TOEVOEGEN PRODUCTIE"
    
    End If
End Sub




Private Sub UserForm_Activate()
ophalen_Productiesoorten

    If Functies.IsUserFormLoaded("FORM_PROJECT_WIJZIGEN") = True Then
        If FORM_PROJECT_WIJZIGEN.productie_inladen = True Then
            Set pr = FORM_PROJECT_WIJZIGEN.prod
            
            'selecteer soort
            For x = 0 To ListBox1.ListCount - 1
                If ListBox1.List(x, 0) = pr.soort Then
                    ListBox1.Selected(x) = True
                    Exit For
                End If
            Next x
            
            'vul de start en einddatum in
            TextBox1 = pr.startdatum
            TextBox2 = pr.einddatum
            FORM_PROJECT_WIJZIGEN.bijwerkenUitvoeringperiode
        End If
    ElseIf Functies.IsUserFormLoaded("FORM_PROJECT_NIEUW") = True Then
        If FORM_PROJECT_NIEUW.productie_inladen = True Then
            Set pr = FORM_PROJECT_NIEUW.producties.item(FORM_PROJECT_NIEUW.ListBox1.ListIndex + 1)
             'selecteer soort
            For x = 0 To ListBox1.ListCount - 1
                If ListBox1.List(x, 0) = pr.soort Then
                    ListBox1.Selected(x) = True
                End If
            Next x
            
            'vul de start en einddatum in
            TextBox1 = pr.startdatum
            TextBox2 = pr.einddatum
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
ophalen_Productiesoorten

    If Functies.IsUserFormLoaded("FORM_PROJECT_WIJZIGEN") = True Then
        If FORM_PROJECT_WIJZIGEN.productie_inladen = True Then
            Set pr = FORM_PROJECT_WIJZIGEN.prod
            
            'selecteer soort
            For x = 0 To ListBox1.ListCount - 1
                If ListBox1.List(x, 0) = pr.soort Then
                    ListBox1.Selected(x) = True
                    Exit For
                End If
            Next x
            
            'vul de start en einddatum in
            TextBox1 = pr.startdatum
            TextBox2 = pr.einddatum
            FORM_PROJECT_WIJZIGEN.bijwerkenUitvoeringperiode
        End If
    ElseIf Functies.IsUserFormLoaded("FORM_PROJECT_NIEUW") = True Then
        If FORM_PROJECT_NIEUW.productie_inladen = True Then
            Set pr = FORM_PROJECT_NIEUW.producties.item(FORM_PROJECT_NIEUW.ListBox1.ListIndex + 1)
             'selecteer soort
            For x = 0 To ListBox1.ListCount - 1
                If ListBox1.List(x, 0) = pr.soort Then
                    ListBox1.Selected(x) = True
                End If
            Next x
            
            'vul de start en einddatum in
            TextBox1 = pr.startdatum
            TextBox2 = pr.einddatum
        End If
    End If
   
End Sub
Function ophalen_Productiesoorten()
Dim lijstproduciesoorten As Variant
Dim db As New DataBase

    
lijstproduciesoorten = db.getLijstBySQL("SELECT * FROM PRODUCTIESOORT")
   
    With ListBox1
        .Clear
        If IsEmpty(lijstproduciesoorten) = False Then
            For r = 0 To UBound(lijstproduciesoorten, 2)
                .AddItem
                .List(.ListCount - 1, 0) = lijstproduciesoorten(0, r)
                .List(.ListCount - 1, 1) = lijstproduciesoorten(1, r)
                .List(.ListCount - 1, 2) = lijstproduciesoorten(2, r)
            Next r
        End If
    End With

End Function
Private Function controleer() As Collection
Set controleer = New Collection

    If TextBox1 = "" Then controleer.Add "Er is geen startdatum gekozen"
    If TextBox2 = "" Then controleer.Add "Er is geen einddatum gekozen"
    If IsDate(TextBox1) = fale Then controleer.Add "De startdatum is geen geldige datum. pas de startdatum aan"
    If IsDate(TextBox2) = fale Then controleer.Add "De einddatum is geen geldige datum. pas de einddatum aan"
    If ListBox1.ListIndex = -1 Then controleer.Add "Er is geen uursoort gekozen"
End Function
Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date

If TextBox1 <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = TextBox1
End If
FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        TextBox1 = ThisWorkbook.datum
    End If
End Sub
Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date

If TextBox2 <> "" Then
    ThisWorkbook.inladen = True
    ThisWorkbook.datum = TextBox2
End If
FORM_KALENDER.Show
    If ThisWorkbook.inladen = True Then
        TextBox2 = ThisWorkbook.datum
    End If
End Sub

