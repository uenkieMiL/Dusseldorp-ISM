VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_MATERIEEL_OVERZICHT 
   Caption         =   "UserForm1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14940
   OleObjectBlob   =   "FORM_MATERIEEL_OVERZICHT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_MATERIEEL_OVERZICHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lijst As Variant
Private lijstBouwjaar As Variant
Private filteractief As Boolean

Private Sub CheckBoxInActief_Click()
    bijwerkenLijstIsActief
End Sub

Private Sub CheckBoxInPlanbaar_Click()
    bijwerkenLijstIsActief
End Sub

Private Sub ComboBoxBouwjaar_Change()
    bijwerkenLijstIsActief
End Sub

Private Sub CommandButton1_Click()
    bijwerkenlijstalles
    filteractief = False
    CheckBoxInActief = False
    CheckBoxInPlanbaar = False

End Sub






Private Sub ListBoxMaterieel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ThisWorkbook.mat_id = ListBoxMaterieel.List(ListBoxMaterieel.ListIndex, 0)
FORM_MATERIEEL.Show
ophalenlijst

bijwerkenlijstalles
End Sub

Private Sub TextBoxFilter_Change()
    If filteractief = False Then
        Bijwerkenlijst
    Else
        bijwerkenLijstIsActief
    End If
End Sub

Private Sub UserForm_Initialize()
ophalenlijst

bijwerkenlijstalles
End Sub

Function ophalenlijst()
Dim db As New DataBase

lijst = db.getLijstBySQL("SELECT * FROM MATERIEEL ORDER BY MaterieelCode")
lijstBouwjaar = db.getLijstBySQL("SELECT DISTINCT Bouwjaar FROM MATERIEEL ORDER BY Bouwjaar DESC")
End Function

Function bijwerkenlijstalles()
Dim r As Long
ListBoxMaterieel.Clear
    If IsEmpty(lijst) = False Then
        For r = 0 To UBound(lijst, 2)
            With ListBoxMaterieel
                .AddItem
                .List(r, 0) = lijst(0, r)
               If IsNull(lijst(1, r)) = False Then .List(r, 1) = lijst(1, r)
                .List(r, 2) = lijst(2, r)
                .List(r, 3) = lijst(3, r)
                .List(r, 4) = lijst(4, r)
                .List(r, 5) = lijst(5, r)
                If IsNull(lijst(6, r)) = False Then .List(r, 6) = lijst(6, r)
                If IsNull(lijst(7, r)) = False Then .List(r, 7) = lijst(7, r)
                If IsNull(lijst(8, r)) = False Then .List(r, 8) = lijst(8, r)
                If IsNull(lijst(9, r)) = False Then .List(r, 9) = lijst(9, r)
            End With
        Next r
    End If
        
    

End Function



Function Bijwerkenlijst()
Dim r As Long
Dim a As Long

ListBoxMaterieel.Clear

If TextBoxFilter = "" Then
    For r = 0 To UBound(lijst, 2)
        With ListBoxMaterieel
            .AddItem
            .List(r, 0) = lijst(0, r)
            If IsNull(lijst(1, r)) = False Then .List(r, 1) = lijst(1, r)
            .List(r, 2) = lijst(2, r)
            .List(r, 3) = lijst(3, r)
            .List(r, 4) = lijst(4, r)
            .List(r, 5) = lijst(5, r)
            If IsNull(lijst(6, r)) = False Then .List(r, 6) = lijst(6, r)
            If IsNull(lijst(7, r)) = False Then .List(r, 7) = lijst(7, r)
            If IsNull(lijst(8, r)) = False Then .List(r, 8) = lijst(8, r)
            If IsNull(lijst(9, r)) = False Then .List(r, 9) = lijst(9, r)
        End With
    Next r
Else
    For r = 0 To UBound(lijst, 2)
        If InStr(1, lijst(1, r), TextBoxFilter, TextCompare) > 0 Or _
        InStr(2, lijst(2, r), TextBoxFilter, TextCompare) > 0 Or _
        InStr(1, lijst(3, r), TextBoxFilter, TextCompare) > 0 Or _
        InStr(1, lijst(4, r), TextBoxFilter, TextCompare) > 0 Or _
        InStr(1, lijst(8, r), TextBoxFilter, TextCompare) > 0 Then
            With ListBoxMaterieel
                .AddItem
                .List(a, 0) = lijst(0, r)
                If IsNull(lijst(1, r)) = False Then .List(a, 1) = lijst(1, r)
                .List(a, 2) = lijst(2, r)
                .List(a, 3) = lijst(3, r)
                .List(a, 4) = lijst(4, r)
                .List(a, 5) = lijst(5, r)
                If IsNull(lijst(6, r)) = False Then .List(a, 6) = lijst(6, r)
                If IsNull(lijst(7, r)) = False Then .List(a, 7) = lijst(7, r)
                If IsNull(lijst(8, r)) = False Then .List(a, 8) = lijst(8, r)
                If IsNull(lijst(9, r)) = False Then .List(a, 9) = lijst(9, r)
            End With
            a = a + 1
        End If
    Next r
 
End If

End Function

Function bijwerkenLijstIsActief()
Dim r As Long
Dim a As Long

ListBoxMaterieel.Clear


If filteractief = False Then filteractief = True

If TextBoxFilter = "" Then
    For r = 0 To UBound(lijst, 2)
            If CheckBoxInActief.Value = lijst(13, r) And CheckBoxInPlanbaar.Value = lijst(12, r) Then
                With ListBoxMaterieel
                    .AddItem
                    .List(a, 0) = lijst(0, r)
                    .List(a, 1) = lijst(1, r)
                    .List(a, 2) = lijst(2, r)
                    .List(a, 3) = lijst(3, r)
                    .List(a, 4) = lijst(4, r)
                    .List(a, 5) = lijst(5, r)
                    .List(a, 6) = lijst(6, r)
                    .List(a, 7) = lijst(7, r)
                    .List(a, 8) = lijst(8, r)
                    .List(a, 9) = lijst(9, r)
                End With
                a = a + 1
            End If
    Next r
Else
     For r = 0 To UBound(lijst, 2)
            If InStr(1, lijst(1, r), TextBoxFilter, TextCompare) > 0 Or _
            InStr(2, lijst(2, r), TextBoxFilter, TextCompare) > 0 Or _
            InStr(1, lijst(3, r), TextBoxFilter, TextCompare) > 0 Or _
            InStr(1, lijst(4, r), TextBoxFilter, TextCompare) > 0 Or _
            InStr(1, lijst(8, r), TextBoxFilter, TextCompare) > 0 Then
                If CheckBoxInActief.Value = lijst(13, r) And CheckBoxInPlanbaar.Value = lijst(12, r) Then
                    With ListBoxMaterieel
                        .AddItem
                        .List(a, 0) = lijst(0, r)
                        .List(a, 1) = lijst(1, r)
                        .List(a, 2) = lijst(2, r)
                        .List(a, 3) = lijst(3, r)
                        .List(a, 4) = lijst(4, r)
                        .List(a, 5) = lijst(5, r)
                        .List(a, 6) = lijst(6, r)
                        .List(a, 7) = lijst(7, r)
                        .List(a, 8) = lijst(8, r)
                        .List(a, 9) = lijst(9, r)
                    End With
                    a = a + 1
                End If
            End If
    Next r
End If

End Function
