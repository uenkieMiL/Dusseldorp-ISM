VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_MATERIEEL 
   Caption         =   "UserForm1"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13290
   OleObjectBlob   =   "FORM_MATERIEEL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_MATERIEEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m As New Materieel


Private Sub CommandButtonBijwerken_Click()
    Dim fouten As Collection
    Dim fout As Variant
    Dim fouttekst As String
    Dim aanpassingen As New Collection
    Dim aanpassing As Variant
    Dim ma As MaterieelActie
    Set fouten = FoutenControle
    
    If fouten.Count = 0 Then
        If m.MaterieelCode <> TextBoxInternNr Then aanpassingen.Add ("Veld Internnummer is aangepast||" & TextBoxInternNr)
        m.MaterieelCode = TextBoxInternNr
        If m.Omschrijving <> TextBoxOmschrijving Then aanpassingen.Add ("Veld Omschrijving is aangepast||" & TextBoxOmschrijving)
        m.Omschrijving = Me.TextBoxOmschrijving
        If m.Omschrijving <> TextBoxOmschrijving Then aanpassingen.Add ("Veld Merk is aangepast||" & TextBoxMerk)
        m.Merk = Me.TextBoxMerk
        If m.MaterieelType <> ComboBoxMateriaalType Then aanpassingen.Add ("Veld Type is aangepast||" & ComboBoxMateriaalType)
        m.MaterieelType = Me.ComboBoxMateriaalType
        If m.Bouwjaar <> TextBoxBouwjaar Then aanpassingen.Add ("Veld Bouwjaar is aangepast||" & TextBoxBouwjaar)
        m.Bouwjaar = Me.TextBoxBouwjaar
        If m.AanschafDatum <> TextBoxAanschafdatum Then aanpassingen.Add ("Veld Aanschafdatum is aangepast||" & TextBoxAanschafdatum)
        m.AanschafDatum = Me.TextBoxAanschafdatum
        If m.KeuringsDatum <> TextBoxKeuringsdatum Then aanpassingen.Add ("Veld KeuringsDatum is aangepast||" & TextBoxKeuringsdatum)
        m.KeuringsDatum = Me.TextBoxKeuringsdatum
        If m.Serienummer <> TextBoxSerienummer Then aanpassingen.Add ("Veld serienummer is aangepast||" & TextBoxSerienummer)
        m.Serienummer = Me.TextBoxSerienummer
        If m.Onderhoudstermijn <> TextBoxOnderhoudstermijn Then aanpassingen.Add ("Veld Onderhoudstermijn is aangepast||" & TextBoxOnderhoudstermijn)
        m.Onderhoudstermijn = Me.TextBoxOnderhoudstermijn
        If Me.TextBoxLaatsteKeuring <> "" Then
            If m.LaatsteOnderhoudsDatum <> TextBoxLaatsteKeuring Then aanpassingen.Add ("Veld Laatste Keuringsdatum is aangepast||" & TextBoxLaatsteKeuring)
            m.LaatsteOnderhoudsDatum = Me.TextBoxLaatsteKeuring
        End If
        If m.Inplanbaar <> CheckBoxInplenbaar Then aanpassingen.Add ("Veld Inplenbaar is aangepast||" & CheckBoxInplenbaar)
        m.Inplanbaar = Me.CheckBoxInplenbaar
        If m.Inactief <> CheckBoxInactief Then aanpassingen.Add ("Veld InActief is aangepast||" & CheckBoxInactief)
        m.Inactief = Me.CheckBoxInactief
        Me.Caption = m.MaterieelCode & " / " & m.Omschrijving
        If Me.LabelId = "" Then
            If m.insert = True Then
                MsgBox "Het materieel is succesvol aangemaakt", vbInformation, "AANMAKEN MATERIEEL"
                Me.LabelId = m.Id
                Me.CommandButtonBijwerken.Caption = "Bijwerken"
                
            Else
                Functies.errorhandler_MsgBox ("Er is iets fout gegaan bij het aanpassen van het materieel")
            End If
        Else
            If m.update = True Then
                MsgBox "Het materieel is succesvol aangepast", vbInformation, "AANPASSEN MATERIEEL"
                For Each aanpassing In aanpassingen
                    Set ma = New MaterieelActie
                    ma.MaterieelId = m.Id
                    ma.InsertBijwerkenVeld CStr(aanpassing)
                Next aanpassing
            Else
                Functies.errorhandler_MsgBox ("Er is iets fout gegaan bij het aanpassen van het materieel")
            End If
        End If
    Else
        fouttekst = "De acties kan niet worden uitgevoerd door de volgende redenen:" & vbNewLine
        
        For Each fout In fouten
            fouttekst = fouttekst & vbNewLine & fout
        Next fout
        Functies.errorhandler_MsgBox (fouttekst)
    End If
    
End Sub

Private Sub CommandButtonBijwerkenFoto_Click()
    m.UpdateFoto
    FotoInladen
End Sub

Private Sub CommandButtonInladen_Click()
    If LabelId <> "" Then
        m.UpdateFoto
        FotoInladen
    End If
End Sub


Private Sub CommandButtonNieuw_Click()
    EmptyForm
End Sub

Private Sub UserForm_Initialize()
    If ThisWorkbook.mat_id <> 0 Then
        m.Id = ThisWorkbook.mat_id
        m.haalop
        bijwerkenForm
    Else
        EmptyForm
    End If
End Sub

Function FotoInladen()
    Image1.Picture = LoadPicture(MateriaalLocatie & m.Foto, 375, 250, Default)
    Image1.PictureSizeMode = fmPictureSizeModeZoom
    
    Me.CommandButtonInladen.Visible = False
    Me.CommandButtonBijwerkenFoto.Visible = True
End Function

Function bijwerkenForm()
    
    Me.TextBoxInternNr = m.MaterieelCode
    Me.TextBoxOmschrijving = m.Omschrijving
    Me.TextBoxMerk = m.Merk
    Me.ComboBoxMateriaalType = m.MaterieelType
    Me.TextBoxBouwjaar = m.Bouwjaar
    Me.TextBoxAanschafdatum = m.AanschafDatum
    Me.TextBoxKeuringsdatum = m.KeuringsDatum
    Me.TextBoxSerienummer = m.Serienummer
    Me.TextBoxOnderhoudstermijn = m.Onderhoudstermijn
    Me.TextBoxLaatsteKeuring = m.LaatsteOnderhoudsDatum
    Me.CheckBoxInplenbaar = m.Inplanbaar
    Me.CheckBoxInactief = m.Inactief
    Me.Caption = m.MaterieelCode & " / " & m.Omschrijving
    Me.LabelId = m.Id
    If m.Foto <> "" Then
        FotoInladen
        Me.LabelFoto = m.Foto
        Else
        Me.CommandButtonBijwerkenFoto.Visible = False
        Me.LabelFoto = ""
    End If
End Function

Function EmptyForm()
    Me.TextBoxInternNr = ""
    Me.TextBoxOmschrijving = ""
    Me.TextBoxMerk = ""
    Me.ComboBoxMateriaalType = ""
    Me.TextBoxBouwjaar = ""
    Me.TextBoxAanschafdatum = ""
    Me.TextBoxKeuringsdatum = ""
    Me.TextBoxSerienummer = ""
    Me.TextBoxOnderhoudstermijn = ""
    Me.TextBoxLaatsteKeuring = ""
    Me.CheckBoxInplenbaar = False
    Me.CheckBoxInactief = False
    Me.Caption = "Materiaal Aanmaken"
    Me.LabelId = ""
    Me.LabelFoto = ""
    Me.CommandButtonBijwerken.Caption = "Aanmaken"
    Me.CommandButtonHistorie.Visible = False
End Function

Function FoutenControle() As Collection
    Dim fout As String
    
    Set FoutenControle = New Collection
    
    If TextBoxInternNr = "" Then
        fout = " - Er is geen Internummer opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxOmschrijving = "" Then
        fout = " - Er is geen omschrijving opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxMerk = "" Then
        fout = " - Er is geen merk opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If ComboBoxMateriaalType = "" Then
        fout = " - Er is geen materiaaltype opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxAanschafdatum = "" Then
        fout = " - Er is geen aanschafdatum opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxKeuringsdatum = "" Then
        fout = " - Er is geen keuringsdatum opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxSerienummer = "" Then
        fout = " - Er is geen serienummer opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If TextBoxOnderhoudstermijn = "" Then
        fout = " - Er is geen onderhoudstermijn opgegegeven. Dit is verplicht"
        FoutenControle.Add fout
    End If
    
    If IsDate(TextBoxAanschafdatum) = False Then
        fout = " - De aanschafdatum blijkt geen geldige datum te zijn. het formaat dient als volgt te zijn: 'dd-mm-jjjj'"
        FoutenControle.Add fout
    End If
    
    If IsDate(TextBoxKeuringsdatum) = False Then
        fout = " - De keuringsdatum blijkt geen geldige datum te zijn. het formaat dient als volgt te zijn: 'dd-mm-jjjj'"
        FoutenControle.Add fout
    End If
    
    If IsDate(TextBoxLaatsteKeuring) = False And TextBoxLaatsteKeuring <> "" Then
        fout = " - De laatste keuringsdatum blijkt geen geldige datum te zijn. het formaat dient als volgt te zijn: 'dd-mm-jjjj'"
        FoutenControle.Add fout
    End If
    
    If IsNumeric(TextBoxBouwjaar) = False Then
        fout = " - het bouwjaar is geen numerieke waarde."
        FoutenControle.Add fout
    End If
    
    If Len(TextBoxBouwjaar) <> 4 Then
        fout = " - het bouwjaar dient een 4 cijferig jaartal te zijn."
        FoutenControle.Add fout
    End If
    
    If TextBoxBouwjaar <> "" Or IsNumeric(TextBoxBouwjaar) = True Then
        If CInt(TextBoxBouwjaar) <= 1980 And CInt(TextBoxBouwjaar) > Year(Now()) Then
            fout = " - het bouwjaar dient tussen 1980 - " & Year(Now()) & " te liggen."
            FoutenControle.Add fout
        End If
    End If
End Function
