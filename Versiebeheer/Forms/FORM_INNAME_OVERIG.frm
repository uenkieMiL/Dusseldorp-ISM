VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_INNAME_OVERIG 
   Caption         =   "INNAME OVERIG MATERIEEL"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "FORM_INNAME_OVERIG.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_INNAME_OVERIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m As New Materieel
Private lijst As Variant
Private lijstMaterieel As Variant


Private Sub CommandButton1_Click()
Dim r As Long
Dim omschr As String
Dim internnr As String
Dim ma As New MaterieelActie
If ListBoxMaterieel.ListIndex > -1 Then
    For r = 0 To ListBoxMaterieel.ListCount - 1
        If ListBoxMaterieel.Selected(r) = True Then
            m.Id = ListBoxMaterieel.List(ListBoxMaterieel.ListIndex, 0)
            m.Status = "In Magazijn"
            If m.updateStatus = False Then
                internnr = ListBoxMaterieel.List(ListBoxMaterieel.ListIndex, 2)
                omschr = ListBoxMaterieel.List(ListBoxMaterieel.ListIndex, 3)
                Functies.errorhandler_MsgBox ("Er is iets misgegaan bij de verwerking van de inname van het volgende materieel" & vbNewLine & _
                                              internnr & " - " & omschr)
            Else
                ma.MaterieelId = m.Id
                ma.Waarde = "In Magazijn"
                ma.InsertInname
                OphalenLijsten
                BijwerkenMaterieelLijst
            End If
            
        End If
    Next r
End If



End Sub



Private Sub TextBoxFilterMaterieel_Change()
BijwerkenMaterieelLijst
End Sub

Private Sub UserForm_Initialize()
OphalenLijsten
BijwerkenMaterieelLijst
End Sub

Function OphalenLijsten()
Dim db As New DataBase
    lijstMaterieel = db.getLijstBySQL("SELECT * FROM MATERIEEL WHERE Inplanbaar = False AND Status <> 'In Magazijn' ORDER BY MaterieelCode;")
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
                .List(r, 1) = lijstMaterieel(4, r) 'Type
                .List(r, 2) = lijstMaterieel(1, r) 'code
                .List(r, 3) = lijstMaterieel(2, r) 'omschrijving
                .List(r, 4) = lijstMaterieel(14, r) 'Status
            End With
        Next r
    Else
       For r = 0 To UBound(lijstMaterieel, 2)
            If InStr(LCase(lijstMaterieel(1, r)), LCase(TextBoxFilterMaterieel)) <> 0 Or InStr(LCase(lijstMaterieel(2, r)), LCase(TextBoxFilterMaterieel)) <> 0 Or InStr(LCase(lijstMaterieel(3, r)), LCase(TextBoxFilterMaterieel)) <> 0 Or InStr(LCase(lijstMaterieel(14, r)), LCase(TextBoxFilterMaterieel)) <> 0 Then
                With ListBoxMaterieel
                    .AddItem
                    .List(a, 0) = lijstMaterieel(0, r) 'id
                    .List(a, 1) = lijstMaterieel(4, r) 'Type
                    .List(a, 2) = lijstMaterieel(1, r) 'code
                    .List(a, 3) = lijstMaterieel(2, r) 'omschrijving
                    .List(a, 4) = lijstMaterieel(14, r) 'Status
                End With
                a = a + 1
            End If
        Next r
    End If
End Function
