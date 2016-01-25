VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_UITGITE_OVERIG 
   Caption         =   "UITGIFTE OVERIG MATERIEEL"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13095
   OleObjectBlob   =   "FORM_UITGITE_OVERIG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_UITGITE_OVERIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private m As New Materieel
Private lijst As Variant
Private lijstMaterieel As Variant


Private Sub CommandButton1_Click()

Dim ma As New MaterieelActie
If ListBoxMaterieel.ListIndex = -1 Then
    Functies.errorhandler_MsgBox ("Er is geen materieel geselecteerd")
    Exit Sub
End If

If TextBoxStatus <> "" And TextBoxStatus <> m.Status Then
    m.Id = CLng(ListBoxMaterieel.List(ListBoxMaterieel.ListIndex, 0))
    m.Status = TextBoxStatus
    If m.updateStatus = True Then
        ma.MaterieelId = m.Id
        ma.Waarde = TextBoxStatus
        ma.InsertUitgifte
        Unload Me
    Else
        Functies.errorhandler_MsgBox ("Er is iets mis gegaan bij het updaten van de status het materieel.")
    End If
End If
End Sub


Private Sub ListBoxProjecten_Click()
    TextBoxStatus = ListBoxProjecten.List(ListBoxProjecten.ListIndex, 0)
End Sub

Private Sub TextBoxFilterMaterieel_Change()
BijwerkenMaterieelLijst
End Sub

Private Sub TextBoxZoeken_Change()
BijwerkenProjectenLijst
End Sub

Private Sub UserForm_Initialize()
OphalenLijsten
BijwerkenProjectenLijst
BijwerkenMaterieelLijst
End Sub

Function OphalenLijsten()
Dim db As New DataBase

    lijst = db.getLijstBySQL("SELECT DISTINCT Synergy, Omschrijving from PROJECTEN ORDER BY Synergy")
    lijstMaterieel = db.getLijstBySQL("SELECT * FROM MATERIEEL WHERE Inplanbaar = False AND Status='In Magazijn' ORDER BY MaterieelCode;")
End Function
    
Function BijwerkenProjectenLijst()
Dim r As Long
Dim a As Long
    ListBoxProjecten.Clear
    
    If TextBoxZoeken = "" Then
        For r = 0 To UBound(lijst, 2)
            With ListBoxProjecten
                .AddItem
                .List(r, 0) = lijst(0, r)
                .List(r, 1) = lijst(1, r)
            End With
        Next r
    Else
       For r = 0 To UBound(lijst, 2)
            If InStr(LCase(lijst(0, r)), LCase(TextBoxZoeken)) <> 0 Or InStr(LCase(lijst(1, r)), LCase(TextBoxZoeken)) <> 0 Then
                With ListBoxProjecten
                    .AddItem
                    .List(a, 0) = lijst(0, r)
                    .List(a, 1) = lijst(1, r)
                End With
                a = a + 1
            End If
        Next r
    End If
End Function

Function BijwerkenMaterieelLijst()
Dim r As Long
Dim a As Long
    ListBoxMaterieel.Clear
    
    If IsEmpty(lijstMaterieel) = True Then Exit Function
        

    
    If TextBoxFilterMaterieel = "" Then
        For r = 0 To UBound(lijstMaterieel, 2)
            With ListBoxMaterieel
                .AddItem
                .List(r, 0) = lijstMaterieel(0, r) 'id
                .List(r, 2) = lijstMaterieel(1, r) 'code
                .List(r, 3) = lijstMaterieel(2, r) 'omschrijving
                .List(r, 1) = lijstMaterieel(3, r) 'Type
            End With
        Next r
    Else
       For r = 0 To UBound(lijstMaterieel, 2)
            If InStr(LCase(lijstMaterieel(1, r)), LCase(TextBoxFilterMaterieel)) <> 0 Or InStr(LCase(lijstMaterieel(2, r)), LCase(TextBoxFilterMaterieel)) <> 0 Or InStr(LCase(lijstMaterieel(3, r)), LCase(TextBoxFilterMaterieel)) <> 0 Then
                With ListBoxMaterieel
                    .AddItem
                    .List(a, 0) = lijstMaterieel(0, r) 'id
                    .List(a, 2) = lijstMaterieel(1, r) 'code
                    .List(a, 3) = lijstMaterieel(2, r) 'omschrijving
                    .List(a, 1) = lijstMaterieel(3, r) 'Type
                End With
                a = a + 1
            End If
        Next r
    End If
End Function
