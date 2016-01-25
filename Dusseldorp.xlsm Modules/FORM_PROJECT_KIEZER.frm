VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PROJECT_KIEZER 
   Caption         =   "UserForm2"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12720
   OleObjectBlob   =   "FORM_PROJECT_KIEZER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PROJECT_KIEZER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private lijstprojecten As Variant


Private Sub CommandButton1_Click()
Dim u As New Uursoort
If TextBoxKleur = "" Then
    MsgBox "Er is geen omschrijving van het uursoort opgegeven. Vul een omschrijving in en druk nogmaals op de plus-knop"
    Exit Sub
End If

If TextBoxKleur.BackColor = -2147483643 Then
    MsgBox "Er is geen achtergrond kleur gekozen. Dubbelklik op het textveld om een achtergrond kleur te kunnen kiezen"
    Exit Sub
End If


u.Kleur = TextBoxKleur.BackColor
u.Omschrijving = TextBoxKleur
If IsNull(CheckBox1) = False Then u.Koppelbaar = True
u.save

ListBox1.AddItem
ListBox1.List(ListBox1.ListCount - 1, 0) = u.Id
ListBox1.List(ListBox1.ListCount - 1, 1) = u.Omschrijving
ListBox1.List(ListBox1.ListCount - 1, 2) = u.Kleur
ListBox1.List(ListBox1.ListCount - 1, 3) = u.Koppelbaar
End Sub

Private Sub CommandButton2_Click()
Dim u As New Uursoort

If ListIndex = -1 Then
MsgBox "Er is geen uursoort geselecteerd om te wijzigen", vbCritical, "FOUT BIJ UURSOORT WIJZIGEN"
Exit Sub
End If

u.Id = CLng(ListBox1.List(ListBox1.ListIndex, 0))
u.Kleur = TextBoxKleur.BackColor
u.Omschrijving = TextBoxKleur
If IsNull(CheckBox1) Then u.Koppelbaar = True Else CheckBox1 = False
u.save

ListBox1.List(ListBox1.ListIndex, 1) = u.Omschrijving
ListBox1.List(ListBox1.ListIndex, 2) = u.Kleur
ListBox1.List(ListBox1.ListIndex, 3) = u.Koppelbaar

TextBoxKleur.BackColor = -2147483643
TextBoxKleur = ""
CheckBox1 = False
ListBox1.Selected(ListBox1.ListIndex) = False
End Sub

Private Sub CommandButton3_Click()
Dim u As New Uursoort
If ListIndex = -1 Then
    MsgBox "Er is geen uursoort geselecteerd om te verwijderen", vbCritical, "FOUT BIJ UURSOORT VERWIJDEREN"
    Exit Sub
End If

u.Id = ListBox1.List(ListBox1.ListIndex, 0)
u.delete

ListBox1.RemoveItem ListBox1.ListIndex
End Sub

Private Sub CommandButton4_Click()
Turbo_AAN
Dim pp As PersoneelPlanning
Dim c As Range
Dim pId As Long
Dim datum As Date
If ListBox2.ListIndex = -1 Then
    If IsNull(CheckBox1) = True And ListBox2.ListIndex = -1 Then
        MsgBox "Er is geen Project gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If
Else
    If ListBox1.ListIndex = -1 Then
        MsgBox "Er is geen Uursoort gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If

End If

For Each c In Selection
        pId = Range("A" & c.Row)
        datum = Cells(1, c.Column)
        c.Interior.Color = TextBoxKleur.BackColor
        If CheckBox1 <> False Then
            c.Value = ListBox2.List(ListBox2.ListIndex, 0)
            c.HorizontalAlignment = xlCenter
            Set pp = New PersoneelPlanning
            pp.datum = datum
            pp.personeelid = pId
            pp.UursoortId = ListBox1.List(ListBox1.ListIndex, 0)
            If IsNull(CheckBox1) = True Then pp.synergy = ListBox2.List(ListBox2.ListIndex, 0)
            pp.save
            If c.Value = "" Then c.Value = pp.synergy Else c.Value = c.Value & Chr(34) & pp.synergy
        End If
Next c
turbo_UIT
Me.Hide
End Sub

Private Sub CommandButton5_Click()
Turbo_AAN
Dim pp As PersoneelPlanning
Dim c As Range
Dim pId As Long
Dim datum As Date
If ListBox2.ListIndex = -1 Then
    If IsNull(CheckBox1) = True And ListBox2.ListIndex = -1 Then
        MsgBox "Er is geen Project gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If
    For Each c In Selection
        pId = Range("A" & c.Row)
        datum = Cells(1, c.Column)
        c.Interior.Color = TextBoxKleur.BackColor
        Set pp = New PersoneelPlanning
        pp.datum = datum
        pp.personeelid = pId
        pp.UursoortId = ListBox1.List(ListBox1.ListIndex, 0)
        pp.save
        If c.Value = "" Then c.Value = UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5)) Else c.Value = c.Value & Chr(10) & Chr(13) & UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5))
    Next c
    
Else
    If ListBox1.ListIndex = -1 Then
        MsgBox "Er is geen Uursoort gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If

  For Each c In Selection
            pId = Range("A" & c.Row)
            datum = Cells(1, c.Column)
            c.Interior.Color = TextBoxKleur.BackColor
            If CheckBox1 <> False Then c.Value = ListBox2.List(ListBox2.ListIndex, 0)
            c.HorizontalAlignment = xlCenter
            Set pp = New PersoneelPlanning
            pp.datum = datum
            pp.personeelid = pId
            pp.UursoortId = ListBox1.List(ListBox1.ListIndex, 0)
            If IsNull(CheckBox1) = True Then pp.synergy = ListBox2.List(ListBox2.ListIndex, 0)
            pp.save
            If c.Value = "" Then c.Value = ListBox2.List(ListBox2.ListIndex, 0) Else c.Value = c.Value & Chr(10) & Chr(13) & pp.synergy
    Next c
    
    With ListBox3
    .AddItem
    .List(.ListCount - 1, 0) = pp.Id
    .List(.ListCount - 1, 1) = ListBox1.List(ListBox1.ListIndex, 0)
    .List(.ListCount - 1, 2) = ListBox1.List(ListBox1.ListIndex, 1)
    If pp.synergy <> "" Then .List(.ListCount - 1, 3) = ListBox2.List(ListBox2.ListIndex, 0)
    End With

End If
turbo_UIT
    Me.Hide
End Sub

Private Sub CommandButton7_Click()
Dim c As Range
Dim pp As PersoneelPlanning

If ListBox3.ListIndex = -1 Then Exit Sub

If ListBox3.List(ListBox3.ListIndex, 3) <> "" Then
    For Each c In Selection
        Set pp = New PersoneelPlanning
        pp.personeelid = Range("A" & c.Row)
        pp.datum = Cells(1, c.Column)
        pp.synergy = ListBox3.List(ListBox3.ListIndex, 3)
        pp.UursoortId = ListBox3.List(ListBox3.ListIndex, 1)
        pp.DeleteDatumPersoneelSynergy
    Next c
Else
    For Each c In Selection
        Set pp = New PersoneelPlanning
        pp.personeelid = Range("A" & c.Row)
        pp.datum = Cells(1, c.Column)
        pp.UursoortId = ListBox3.List(ListBox3.ListIndex, 1)
        pp.DeleteDatumPersoneelUursoort
    Next c
End If

    
Me.Hide
PersoneelsPlanning.MaakPersoneelsPlanning
End Sub

Private Sub ListBox1_Click()
TextBoxKleur = ListBox1.List(ListBox1.ListIndex, 1)
TextBoxKleur.BackColor = ListBox1.List(ListBox1.ListIndex, 2)
CheckBox1 = ListBox1.List(ListBox1.ListIndex, 3)
End Sub

Function getLijstDag(Id As Long, datum As Date) As Collection
Dim lijstvar As Variant
Dim db As New DataBase

Dim pp As PersoneelPlanning
Set getLijstDag = New Collection
lijstvar = db.getLijstBySQL("SELECT PLANNING_PERSONEEL.*, UURSOORT.* FROM PLANNING_PERSONEEL INNER JOIN UURSOORT ON PLANNING_PERSONEEL.UursoortId = Uursoort.Id WHERE PersoneelId = " _
& Id & " AND Datum = #" & Month(datum) & "/" & Day(datum) & "/" & Year(datum) & "#")
If IsEmpty(lijstvar) = False Then
    For l = 0 To UBound(lijstvar, 2)
    Set pp = New PersoneelPlanning
    pp.Id = lijstvar(0, l)
    pp.UursoortId = lijstvar(3, l)
    pp.synergy = lijstvar(4, l)
    pp.Uursoort.Id = lijstvar(5, l)
    pp.Uursoort.Omschrijving = lijstvar(6, l)
    getLijstDag.Add pp, CStr(pp.Id)
    Next l
End If
End Function

Private Sub TextBox1_Change()

ListBox2.Clear
If TextBox1 = "" Then
    For x = 0 To UBound(lijstprojecten, 2)
        With ListBox2
            .AddItem
            .List(x, 0) = lijstprojecten(0, x)
            .List(x, 1) = lijstprojecten(1, x)
        End With
    Next x
Else
     For x = 0 To UBound(lijstprojecten, 2)
        If InStr(LCase(lijstprojecten(0, x)), LCase(TextBox1)) <> 0 Or InStr(LCase(lijstprojecten(1, x)), LCase(TextBox1)) <> 0 Then
            With ListBox2
                .AddItem
                .List(a, 0) = lijstprojecten(0, x)
                .List(a, 1) = lijstprojecten(1, x)
                a = a + 1
            End With
        End If
    Next x
End If

End Sub

Private Sub TextBoxKleur_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If TextBoxKleur.BackColor = -2147483643 Then
    TextBoxKleur.BackColor = PickNewColor
Else
    TextBoxKleur.BackColor = PickNewColor(TextBoxKleur.BackColor)
End If
End Sub


Private Sub UserForm_Activate()
bijwerkenLijstDag
End Sub
Function bijwerkenLijstDag()
Dim lijstdag As New Collection
Dim datum As Date
Dim Id As Long

Id = ThisWorkbook.Sheets(Blad5.Name).Range(col_pers_id & ActiveCell.Row)
datum = ThisWorkbook.Sheets(Blad5.Name).Range(Cells(PersoneelsPlanning.row_datum, ActiveCell.Column).Address)

Set lijstdag = getLijstDag(Id, datum)

ListBox3.Clear

For Each pp In lijstdag
    With ListBox3
        .AddItem
        .List(x, 0) = pp.Id
        .List(x, 1) = pp.Uursoort.Id
        .List(x, 2) = pp.Uursoort.Omschrijving
        .List(x, 3) = pp.synergy
    End With
    x = x + 1
Next pp
End Function

Private Sub UserForm_Initialize()
Dim lijst As Variant
Dim db As New DataBase

Dim pp As PersoneelPlanning
Dim datum As Date
Dim Id As Long
Dim titel As String
datum = CDate(Cells(PersoneelsPlanning.row_datum, ActiveCell.Column))
titel = ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_bedrijf & ActiveCell.Row)
titel = titel & " / " & ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_voornaam & ActiveCell.Row)
titel = titel & " " & ThisWorkbook.Sheets(Blad5.Name).Range(PersoneelsPlanning.col_pers_achternaam & ActiveCell.Row)
titel = titel & " / " & FormatDateTime(datum, vbShortDate)
FORM_PROJECT_KIEZER.Caption = titel

lijst = db.getLijstBySQL("select * from UURSOORT WHERE InActief = False")
lijstprojecten = db.getLijstBySQL("SELECT DISTINCT PLANNINGEN.Synergy, PROJECTEN.Omschrijving " & _
"FROM PROJECTEN INNER JOIN PLANNINGEN ON PROJECTEN.Synergy = PLANNINGEN.Synergy " & _
"WHERE (((PLANNINGEN.Soort)=4) AND ((PLANNINGEN.STATUS)=False)) ORDER BY PLANNINGEN.Synergy;")

If IsEmpty(lijst) = False Then
    For x = 0 To UBound(lijst, 2)
        With ListBox1
            .AddItem
            .List(x, 0) = lijst(0, x)
            .List(x, 1) = lijst(1, x)
            .List(x, 2) = lijst(2, x)
            .List(x, 3) = lijst(3, x)
        End With
    Next x
End If

If IsEmpty(lijstprojecten) = False Then
    For x = 0 To UBound(lijstprojecten, 2)
        With ListBox2
            .AddItem
            .List(x, 0) = lijstprojecten(0, x)
            .List(x, 1) = lijstprojecten(1, x)
        End With
    Next x
End If

bijwerkenLijstDag


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
turbo_UIT
End Sub
