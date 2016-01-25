VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_UITGITE_GEPLAND 
   Caption         =   "UserForm1"
   ClientHeight    =   3390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   OleObjectBlob   =   "FORM_UITGITE_GEPLAND.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_UITGITE_GEPLAND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m As New Materieel
Private lijst As Variant


Private Sub CommandButton1_Click()
Dim ma As New MaterieelActie
If TextBoxStatus <> "" And TextBoxStatus <> m.Status Then
    m.Status = TextBoxStatus
    If m.updateStatus = True Then
        ThisWorkbook.Sheets(Blad4.Name).Range(MaterielenPlanning.col_mat_status & ActiveCell.Row) = TextBoxStatus
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

Private Sub TextBoxZoeken_Change()
BijwerkenProjectenLijst
End Sub

Private Sub UserForm_Initialize()

If ThisWorkbook.mat_id <> 0 Then
    m.Id = ThisWorkbook.mat_id
    m.haalop
    TextBoxStatus = m.Status
    
    Me.Caption = m.MaterieelCode & " / " & m.Omschrijving
End If



OphalenProjectenlijst
BijwerkenProjectenLijst

End Sub

Function OphalenProjectenlijst()
Dim db As New DataBase

    lijst = db.getLijstBySQL("SELECT DISTINCT Synergy, Omschrijving from PROJECTEN ORDER BY Synergy")
    
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
