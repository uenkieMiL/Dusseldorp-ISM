VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_PRODUCTIE_AANMAKEN 
   Caption         =   "Productie Aanmaken"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4425
   OleObjectBlob   =   "FORM_PRODUCTIE_AANMAKEN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_PRODUCTIE_AANMAKEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
Dim pr As New Productie
Dim Controle As New Collection
Dim tekst As String
Dim fout As Variant

    Set Controle = controleer
    
    If Controle.Count = 0 Then
        If Functies.IsUserFormLoaded("FORM_PROJECT_WIJZIGEN") = True Then
            pr.synergy = FORM_PROJECT_WIJZIGEN.TextBox1
            pr.Vestiging = FORM_PROJECT_WIJZIGEN.Combo_Vestiging
            pr.soort = ListBox1.List(ListBox1.ListIndex, 0)
            pr.Kleur = ListBox1.List(ListBox1.ListIndex, 2)
            pr.startdatum = CDate(TextBox1)
            pr.einddatum = CDate(TextBox2)
            pr.insert
            pr.Omschrijving = ListBox1.List(ListBox1.ListIndex, 1)
            FORM_PROJECT_WIJZIGEN.project.ToevoegenProductie pr
            FORM_PROJECT_WIJZIGEN.newproductie = True
            Me.Hide
        ElseIf Functies.IsUserFormLoaded("FORM_PROJECT_AANMAKEN") = True Then
            pr.soort = ListBox1.List(ListBox1.ListIndex, 0)
            pr.Omschrijving = ListBox1.List(ListBox1.ListIndex, 1)
            pr.Kleur = ListBox1.List(ListBox1.ListIndex, 2)
            pr.startdatum = CDate(TextBox1)
            pr.einddatum = CDate(TextBox2)
            pr.Omschrijving = ListBox1.List(ListBox1.ListIndex, 1)
            FORM_PROJECT_AANMAKEN.producties.Add pr
            FORM_PROJECT_AANMAKEN.newproductie = True
            Me.Hide
        End If
    Else
    'fouten. informeer gebruiker
        tekst = "De Productie kan niet worden toegevoegd aan het project om de volgende reden(en):"
        For Each fout In Controle
        tekst = tekst & vbNewLine & fout
        Next fout
    MsgBox tekst, vbCritical, "FOUT BIJ TOEVOEGEN PRODUCTIE"
    
    End If

End Sub


Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim startdatum As Date

If Combo_Acquisitie_Start <> "" Then
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

Private Sub UserForm_Initialize()
    ophalen_Productiesoorten
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

Function getPlanningID(synergy As String, soort As Byte) As Long
Dim lijst As Variant
Dim sql As String

sql = "SELECT TOP 1 Id FROM PLANNINGEN WHERE synergy = '" & synergy & "' AND SOORT = " & soort
lijst = DataBase.LijstOpBasisVanQuery(sql)

getPlanningID = CLng(lijst(0, 0))
End Function

Private Function controleer() As Collection
Set controleer = New Collection

    If TextBox1 = "" Then controleer.Add "Er is geen startdatum gekozen"
    If TextBox2 = "" Then controleer.Add "Er is geen einddatum gekozen"
    If IsDate(TextBox1) = fale Then controleer.Add "De startdatum is geen geldige datum. pas de startdatum aan"
    If IsDate(TextBox2) = fale Then controleer.Add "De einddatum is geen geldige datum. pas de einddatum aan"
    If ListBox1.ListIndex = -1 Then controleer.Add "Er is geen uursoort gekozen"
End Function

