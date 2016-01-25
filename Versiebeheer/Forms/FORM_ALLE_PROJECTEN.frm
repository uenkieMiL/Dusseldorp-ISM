VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_ALLE_PROJECTEN 
   Caption         =   "PROJECTEN OVERZICHT"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12210
   OleObjectBlob   =   "FORM_ALLE_PROJECTEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_ALLE_PROJECTEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim lijst As Variant





Private Sub CommandButton1_Click()
    Dim p As New project
    Dim synergy As String
    Dim Vestiging As String
    
    If ListBox1.ListIndex > -1 Then
        synergy = ListBox1.List(ListBox1.ListIndex, 0)
        Vestiging = ListBox1.List(ListBox1.ListIndex, 1)
        antwoord = MsgBox("Weet u zeker dat u project " & synergy & " wilt vewijderen?", vbYesNo, "VERWIJDEREN PROJECT")
        If antwoord = vbYes Then
            p.synergy = synergy
            p.Vestiging = Vestiging
            p.haalop
            p.verwijderenProject
            ListBox1.RemoveItem ListBox1.ListIndex
        End If
    End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim synergy As String

ThisWorkbook.synergy_id = ListBox1.List(ListBox1.ListIndex, 0)
ThisWorkbook.Vestiging = ListBox1.List(ListBox1.ListIndex, 1)
FORM_PROJECT_WIJZIGEN.Show
TextBox1 = ""
End Sub

Private Sub OptionButton1_Click()
TextBox1 = ""
lijst = DataBase.LijstOpBasisVanQuery("SELECT * FROM projecten")
    ListBox1.Clear
    For x = 0 To UBound(lijst, 2)
        With ListBox1
            .AddItem
            .List(x, 0) = lijst(0, x)
                .List(x, 1) = lijst(10, x)
                .List(x, 2) = lijst(1, x)
                .List(x, 3) = lijst(2, x)
        End With
    Next x
End Sub

Private Sub OptionButton2_Click()
TextBox1 = ""
lijst = DataBase.LijstCalculatieNiet
    ListBox1.Clear
    If IsEmpty(lijst) = False Then
        For x = 0 To UBound(lijst, 2)
            With ListBox1
                .AddItem
                .List(x, 0) = lijst(0, x)
                .List(x, 1) = lijst(10, x)
                .List(x, 2) = lijst(1, x)
                .List(x, 3) = lijst(2, x)
            End With
        Next x
    End If
End Sub

Private Sub OptionButton3_Click()
lijst = DataBase.LijstProjectenAfgerond
    ListBox1.Clear
    If IsEmpty(lijst) = False Then
        For x = 0 To UBound(lijst, 2)
            With ListBox1
                .AddItem
                .List(x, 0) = lijst(0, x)
                .List(x, 1) = lijst(10, x)
                .List(x, 2) = lijst(1, x)
                .List(x, 3) = lijst(2, x)
            End With
        Next x
    End If
End Sub


Private Sub TextBox1_Change()
    Dim a As Long
    If IsEmpty(lijst) = True Then Exit Sub
    ListBox1.Clear
    If TextBox1 = "" Then
        For x = 0 To UBound(lijst, 2)
            With ListBox1
                .AddItem
                .List(x, 0) = lijst(0, x)
                .List(x, 1) = lijst(10, x)
                .List(x, 2) = lijst(1, x)
                .List(x, 3) = lijst(2, x)
            End With
        Next x
    Else
        For x = 0 To UBound(lijst, 2)
            If InStr(lijst(0, x), TextBox1) <> 0 Or InStr(LCase(lijst(1, x)), LCase(TextBox1)) <> 0 Or InStr(LCase(lijst(2, x)), LCase(TextBox1)) <> 0 Then
    
                With ListBox1
                .AddItem
                .List(a, 0) = lijst(0, x)
                .List(a, 1) = lijst(10, x)
                .List(a, 2) = lijst(1, x)
                .List(a, 3) = lijst(2, x)
                End With
                a = a + 1
            End If
        Next x
    End If
End Sub

Private Sub UserForm_Initialize()
Dim db As New DataBase

lijst = db.getLijstBySQL("SELECT * FROM projecten ORDER BY Synergy")

For x = 0 To UBound(lijst, 2)
    With ListBox1
        .AddItem
        .List(x, 0) = lijst(0, x)
        .List(x, 1) = lijst(10, x)
        .List(x, 2) = lijst(1, x)
        .List(x, 3) = lijst(2, x)
    End With
Next x
End Sub
