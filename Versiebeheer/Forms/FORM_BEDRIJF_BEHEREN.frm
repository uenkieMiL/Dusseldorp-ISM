VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_BEDRIJF_BEHEREN 
   Caption         =   "UserForm1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13260
   OleObjectBlob   =   "FORM_BEDRIJF_BEHEREN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_BEDRIJF_BEHEREN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim lijst As Collection

Private Sub CommandButton1_Click()
If TextBoxBedrijfsnaam = "" Or TextBoxKVK = "" Then
    MsgBox "Er is geen Bedrijfsnaam of KVK nummer ingevoerd. AUB beide velden invullen.", vbCritical, "FOUT BIJ AANMAKEN BEDRIJF"
    Exit Sub
End If

Dim b As New Bedrijf
b.Bedrijfsnaam = TextBoxBedrijfsnaam
b.Contactpersoon = TextBoxContactpersoon
b.KVK = TextBoxKVK
b.Telefoonnummer = TextBoxTel
b.Contactpersoon = TextBoxContactpersoon
b.Emailadres = TextboxMail
b.save

With ListBox1
.AddItem
.List(.ListCount - 1, 0) = b.Id
.List(.ListCount - 1, 1) = b.Bedrijfsnaam
.List(.ListCount - 1, 2) = b.KVK
.List(.ListCount - 1, 3) = b.Contactpersoon
.List(.ListCount - 1, 4) = b.Telefoonnummer
.List(.ListCount - 1, 5) = b.Emailadres
End With
End Sub

Private Sub CommandButton2_Click()
Dim b As New Bedrijf

If ListBox1.ListIndex = -1 Then
    MsgBox "Er is geen bedrijf geselecteerd", vbCritical, "GEEN BEDRIJF GESELECTEERD"
    Exit Sub
End If

b.Id = ListBox1.List(ListBox1.ListIndex, 0)
b.Bedrijfsnaam = TextBoxBedrijfsnaam
b.KVK = TextBoxKVK
b.Contactpersoon = TextBoxContactpersoon
b.Telefoonnummer = TextBoxTel
b.Emailadres = TextboxMail
b.save


ListBox1.List(ListBox1.ListIndex, 1) = b.Bedrijfsnaam
ListBox1.List(ListBox1.ListIndex, 2) = b.KVK
ListBox1.List(ListBox1.ListIndex, 3) = b.Contactpersoon
ListBox1.List(ListBox1.ListIndex, 4) = b.Telefoonnummer
ListBox1.List(ListBox1.ListIndex, 5) = b.Emailadres
End Sub

Private Sub CommandButton3_Click()
leegvelden
End Sub

Private Sub CommandButton4_Click()
Dim b As New Bedrijf

If ListBox1.ListIndex <> -1 Then
    b.Id = ListBox1.List(ListBox1.ListIndex, 0)
    b.Bedrijfsnaam = ListBox1.List(ListBox1.ListIndex, 1)
    b.KVK = ListBox1.List(ListBox1.ListIndex, 2)
    b.Contactpersoon = ListBox1.List(ListBox1.ListIndex, 3)
    b.Telefoonnummer = ListBox1.List(ListBox1.ListIndex, 4)
    b.Emailadres = ListBox1.List(ListBox1.ListIndex, 5)
    
    Set FORM_PERSONEEL_OVERZICHT.b = b
    FORM_PERSONEEL_OVERZICHT.TextBoxBedrijf = b.Bedrijfsnaam
    FORM_PERSONEEL_OVERZICHT.TextBoxBedrijfId = b.Id
    Me.Hide
End If

End Sub

Private Sub ListBox1_Click()
TextBoxBedrijfsnaam = ListBox1.List(ListBox1.ListIndex, 1)
TextBoxKVK = ListBox1.List(ListBox1.ListIndex, 2)
TextBoxContactpersoon = ListBox1.List(ListBox1.ListIndex, 3)
TextBoxTel = ListBox1.List(ListBox1.ListIndex, 4)
TextboxMail = ListBox1.List(ListBox1.ListIndex, 5)
End Sub

Private Sub TextBox1_Change()
    Dim a As Long
    Dim r As Long
    Dim b As Bedrijf
    If lijst.Count = 0 Then Exit Sub
    ListBox1.Clear
    If TextBox1 = "" Then
        For Each b In lijst
            With ListBox1
                .AddItem
                .List(r, 0) = b.Id
                .List(r, 1) = b.Bedrijfsnaam
                .List(r, 2) = b.KVK
                .List(r, 3) = b.Contactpersoon
                .List(r, 4) = b.Telefoonnummer
                .List(r, 5) = b.Emailadres
                r = r + 1
            End With
        Next b
    Else
        For Each b In lijst
            If InStr(b.Id, TextBox1) <> 0 Or InStr(LCase(b.Bedrijfsnaam), LCase(TextBox1)) <> 0 Or InStr(LCase(b.KVK), LCase(TextBox1)) <> 0 Or InStr(LCase(b.Contactpersoon), LCase(TextBox1)) <> 0 Or InStr(LCase(b.Telefoonnummer), LCase(TextBox1)) <> 0 Or InStr(LCase(b.Emailadres), LCase(TextBox1)) <> 0 Then
    
                With ListBox1
                .AddItem
                .List(a, 0) = b.Id
                .List(a, 1) = b.Bedrijfsnaam
                .List(a, 2) = b.KVK
                .List(a, 3) = b.Contactpersoon
                .List(a, 4) = b.Telefoonnummer
                .List(a, 5) = b.Emailadres
                End With
                a = a + 1
            End If
        Next b
    End If
End Sub



Private Sub TextBoxBedrijfsnaam_Change()
If TextBoxBedrijfsnaam = "" Then
    If ListBox1.ListIndex > -1 Then CommandButton1.Visible = False
    CommandButton2.Visible = False
    CommandButton3.Visible = False
Else
    CommandButton1.Visible = True
    CommandButton2.Visible = True
    CommandButton3.Visible = True
End If
End Sub

Private Sub UserForm_Initialize()

Dim b As Bedrijf
Dim r As Long

    CommandButton1.Visible = False
    CommandButton2.Visible = False
    CommandButton3.Visible = False

    
Set lijst = New Collection
Set lijst = getBedrijvenLijst

For Each b In lijst
    With ListBox1
    .AddItem
        .List(r, 0) = b.Id
        .List(r, 1) = b.Bedrijfsnaam
        .List(r, 2) = b.KVK
        .List(r, 3) = b.Contactpersoon
        .List(r, 4) = b.Telefoonnummer
        .List(r, 5) = b.Emailadres
    End With
    r = r + 1
Next b
End Sub



Function getBedrijvenLijst() As Collection
Dim lijst As Variant
Dim b As Bedrijf
Dim db As New DataBase
Set getBedrijvenLijst = New Collection

lijst = db.getLijstBySQL("SELECT * FROM BEDRIJVEN ORDER BY Bedrijfsnaam")

For x = 0 To UBound(lijst, 2)
    Set b = New Bedrijf
    b.Id = lijst(0, x)
    b.KVK = lijst(1, x)
    b.Bedrijfsnaam = lijst(2, x)
    If IsNull(lijst(3, x)) = False Then b.Contactpersoon = lijst(3, x)
    If IsNull(lijst(4, x)) = False Then b.Telefoonnummer = lijst(4, x)
    If IsNull(lijst(5, x)) = False Then b.Emailadres = lijst(5, x)
    getBedrijvenLijst.Add b, CStr(b.Id)
    
    
Next x

End Function

Function leegvelden()
TextBoxBedrijfsnaam = ""
TextBoxKVK = ""
TextBoxContactpersoon = ""
TextBoxTel = ""
TextboxMail = ""

If ListBox1.ListIndex <> -1 Then ListBox1.Selected(ListBox1.ListIndex) = False

End Function
