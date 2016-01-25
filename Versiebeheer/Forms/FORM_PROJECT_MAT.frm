VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PROJECT_MAT 
   Caption         =   "UserForm2"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   OleObjectBlob   =   "FORM_PROJECT_MAT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PROJECT_MAT"
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
Dim mp As MaterieelPlanning
Dim c As Range
Dim pId As Long
Dim datum As Date
If ListBox2.ListIndex = -1 Then
    If IsNull(CheckBox1) = True And ListBox2.ListIndex = -1 Then
        MsgBox "Er is geen Project gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If
    
    If IsDate(TextBoxStartdatum) = False Or IsDate(TextBoxEinddatum) = False Then
        MsgBox "Er is geen correct start en einddatum gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If
    
    For Each c In Selection
        pId = Range("A" & c.Row)
        datum = Cells(1, c.Column)
        If c.Column >= MaterielenPlanning.startkolom Then
            c.Interior.Color = TextBoxKleur.BackColor
            If c.Value = "" Then c.Value = UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5)) Else c.Value = c.Value & Chr(10) & Chr(13) & UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5))
        End If
    Next c
    
        Set mp = New MaterieelPlanning
        mp.startdatum = CDate(TextBoxStartdatum)
        mp.einddatum = CDate(TextBoxEinddatum)
        mp.MaterieelId = pId
        mp.MaterieelSoortId = ListBox1.List(ListBox1.ListIndex, 0)
        If mp.insert = True Then
        
        With ListBoxGepland
            .AddItem
            .List(.ListCount - 1, 0) = mp.Id
            .List(.ListCount - 1, 1) = ListBox1.List(ListBox1.ListIndex, 0)
            .List(.ListCount - 1, 2) = ListBox1.List(ListBox1.ListIndex, 1)
            If mp.synergy <> "" Then .List(.ListCount - 1, 3) = ListBox2.List(ListBox2.ListIndex, 0)
            .List(.ListCount - 1, 4) = FormatDateTime(TextBoxStartdatum, vbShortDate)
            .List(.ListCount - 1, 5) = FormatDateTime(TextBoxEinddatum, vbShortDate)
        End With
    End If
Else
    If ListBox1.ListIndex = -1 Then
        MsgBox "Er is geen Materieelsoort gekozen", vbCritical, "FOUT BIJ INPLANNEN"
        turbo_UIT
        Exit Sub
    End If

    For Each c In Selection
            If c.Column >= MaterielenPlanning.startkolom Then
                pId = Range("A" & c.Row)
                datum = Cells(1, c.Column)
                
                c.Interior.Color = TextBoxKleur.BackColor
                If CheckBox1 <> False Then c.Value = ListBox2.List(ListBox2.ListIndex, 0)
                c.HorizontalAlignment = xlCenter
                If IsNull(CheckBox1) = True Then
                    If c.Value = "" Then c.Value = ListBox2.List(ListBox2.ListIndex, 0) Else c.Value = c.Value & Chr(10) & Chr(13) & ListBox2.List(ListBox2.ListIndex, 0)
                Else
                    If c.Value = "" Then c.Value = UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5)) Else c.Value = c.Value & Chr(10) & Chr(13) & UCase(Left(ListBox1.List(ListBox1.ListIndex, 1), 5))
                End If
            End If
    Next c
    
        Set mp = New MaterieelPlanning
        mp.startdatum = CDate(TextBoxStartdatum)
        mp.einddatum = CDate(TextBoxEinddatum)
        mp.MaterieelId = pId
        mp.MaterieelSoortId = ListBox1.List(ListBox1.ListIndex, 0)
        If IsNull(CheckBox1) = True Or CheckBox1 = True Then
            mp.synergy = ListBox2.List(ListBox2.ListIndex, 0)
            mp.Gekoppeld = True
        End If
     If mp.insert = True Then

        With ListBoxGepland
        .AddItem
        .List(.ListCount - 1, 0) = mp.Id
        .List(.ListCount - 1, 1) = ListBox1.List(ListBox1.ListIndex, 0)
        .List(.ListCount - 1, 2) = ListBox1.List(ListBox1.ListIndex, 1)
        If mp.synergy <> "" Then .List(.ListCount - 1, 3) = ListBox2.List(ListBox2.ListIndex, 0)
        .List(.ListCount - 1, 4) = TextBoxStartdatum
        .List(.ListCount - 1, 5) = TextBoxEinddatum
        
        End With
    End If

End If


turbo_UIT
    Me.Hide
End Sub

Private Function VoegMaterieelPlanningAanGepland(mt As MaterieelPlanning)

End Function

Private Sub CommandButton7_Click()
'verwijderne planning
Dim c As Range
Dim mp As MaterieelPlanning

If ListBoxGepland.ListIndex = -1 Then Exit Sub

        Set mp = New MaterieelPlanning
        mp.Id = ListBoxGepland.List(ListBoxGepland.ListIndex, 0)
        If mp.delete = True Then ListBoxGepland.RemoveItem (ListBoxGepland.ListIndex)
    
Me.Hide
MaterielenPlanning.MaterieelPlanningVernieuwen
End Sub

Private Sub ListBox1_Click()
TextBoxKleur = ListBox1.List(ListBox1.ListIndex, 1)
TextBoxKleur.BackColor = ListBox1.List(ListBox1.ListIndex, 2)
CheckBox1 = ListBox1.List(ListBox1.ListIndex, 3)
End Sub

Function getLijstDag(Id As Long, datum As Date) As Collection
Dim lijstvar As Variant
Dim db As New DataBase

Dim mp As MaterieelPlanning
Set getLijstDag = New Collection
lijstvar = db.getLijstBySQL("SELECT PLANNING_MATERIEEL.*, MATERIEELSOORT.* FROM PLANNING_MATERIEEL INNER JOIN MATERIEELSOORT ON PLANNING_MATERIEEL.MaterieelSoortId = Materieelsoort.Id WHERE MaterieelId = " _
& Id & " AND (PLANNING_MATERIEEL.StartDatum >= #" & Month(datum) & "/" & Day(datum) & "/" & Year(datum) & "# OR PLANNING_MATERIEEL.EindDatum >= #" & Month(datum) & "/" & Day(datum) & "/" & Year(datum) & "#) ORDER BY PLANNING_MATERIEEL.StartDatum;")
If IsEmpty(lijstvar) = False Then
    For l = 0 To UBound(lijstvar, 2)
    Set mp = New MaterieelPlanning
    mp.Id = lijstvar(0, l)
    mp.MaterieelId = lijstvar(1, l)
    mp.startdatum = lijstvar(2, l)
    mp.einddatum = lijstvar(3, l)
    mp.synergy = lijstvar(6, l)
    mp.MaterieelSoortId = lijstvar(8, l)
    mp.Materieel.Omschrijving = lijstvar(9, l)
    getLijstDag.Add mp, CStr(mp.Id)
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
Dim mp As MaterieelPlanning
Id = ThisWorkbook.Sheets(Blad4.Name).Range(col_mat_id & ActiveCell.Row)
datum = ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_datum, MaterielenPlanning.col_plan_start).Address)

Set lijstdag = getLijstDag(Id, datum)

ListBoxGepland.Clear

For Each mp In lijstdag
    With ListBoxGepland
        .AddItem
        .List(x, 0) = mp.Id
        .List(x, 1) = mp.MaterieelSoortId
        .List(x, 2) = mp.Materieel.Omschrijving
        .List(x, 3) = mp.synergy
        .List(x, 4) = mp.startdatum
        .List(x, 5) = mp.einddatum
    End With
    x = x + 1
Next mp
End Function

Private Sub UserForm_Initialize()
Dim lijst As Variant
Dim db As New DataBase

Dim mp As MaterieelPlanning
Dim startdatum As Date
Dim einddatum As Date
Dim Id As Long
Dim titel As String
Dim c As Range
datum = CDate(Cells(MaterielenPlanning.row_datum, ActiveCell.Column))
titel = ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_intern & ActiveCell.Row)
titel = titel & " / " & ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_omschr & ActiveCell.Row)
titel = titel & " / " & FormatDateTime(datum, vbShortDate)
FORM_PROJECT_MAT.Caption = titel

If startdatum = #12:00:00 AM# Then startdatum = ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_datum, Selection.Column).Address).Value
If einddatum = #12:00:00 AM# Then einddatum = ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_datum, Selection.Column + Selection.Columns.Count - 1).Address).Value
If startdatum = #12:00:00 AM# And einddatum <> #12:00:00 AM# Then startdatum = ThisWorkbook.Sheets(Blad4.Name).Range(Cells(MaterielenPlanning.row_datum, MaterielenPlanning.col_plan_start).Address).Value

If startdatum <> #12:00:00 AM# Then TextBoxStartdatum = FormatDateTime(startdatum, vbShortDate)
If einddatum <> #12:00:00 AM# Then TextBoxEinddatum = FormatDateTime(einddatum, vbShortDate)

lijst = db.getLijstBySQL("select * from MATERIEELSOORT WHERE InActief = False")
lijstprojecten = db.getLijstBySQL("SELECT DISTINCT Synergy, Omschrijving " & _
"FROM PROJECTEN GROUP BY Synergy, Omschrijving;")

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
