VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_STAATINWACHT 
   Caption         =   "PROJECTEN IN DE WACHT"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16365
   OleObjectBlob   =   "FORM_STAATINWACHT.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_STAATINWACHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private overzicht As Variant
Private lijst As Variant

Private Sub ToggleButton1_Click()
UpdateProject
MaakArchiefLijst
End Sub

Private Sub CommandButton1_Click()
Dim rng As Range
Dim ws As String
ws = ActiveSheet.Name
Set rng = ActiveCell
UpdateProject

If ws = Blad1.Name Then
    SoortPlanning.MaakSoortPlanning ws, 1
ElseIf ws = Blad2.Name Then
    SoortPlanning.MaakSoortPlanning ws, 2
ElseIf ws = Blad3.Name Then
    SoortPlanning.MaakSoortPlanning ws, 4
End If
Unload Me
rng.Select

End Sub

Private Sub CommandButton2_Click()
Dim datumtekst As String
Dim synergy As String
Dim Vestiging As String
Dim datum As Date
Dim p As New project

If ListBox1.ListIndex = -1 Then Exit Sub

synergy = ListBox1.List(ListBox1.ListIndex, 0)
Vestiging = ListBox1.List(ListBox1.ListIndex, 3)
If ListBox1.List(ListBox1.ListIndex, 4) = "" Then ListBox1.List(ListBox1.ListIndex, 4) = FormatDateTime(Now(), vbShortDate)
ThisWorkbook.datum = ListBox1.List(ListBox1.ListIndex, 5)
ThisWorkbook.inladen = True
FORM_KALENDER.Show

If synergy <> "" And IsNumeric(synergy) = True And ThisWorkbook.inladen = True Then
        p.synergy = synergy
        p.Vestiging = Vestiging
        p.staatInWacht = True
        p.naBelDatum = ThisWorkbook.datum
        If p.UpdateWacht = True Then MaakArchiefLijst
    Else
        MsgBox "Er is geen datum opgegeven, deze actie wordt geanuleerd" & vbNewLine & "Er is ingevoerd : " & datumtekst, vbCritical, "DATUM ONJUIST"
    End If

End Sub

Private Sub CommandButton3_Click()
    Dim p As New project
    Dim synergy As String
    Dim Vestiging As String
    synergy = ListBox1.List(ListBox1.ListIndex, 3)
    If ListBox1.ListIndex > -1 Then
        synergy = ListBox1.List(ListBox1.ListIndex, 0)
        Vestiging = ListBox1.List(ListBox1.ListIndex, 3)
        antwoord = MsgBox("Weet u zeker dat u project " & synergy & " wilt vewijderen?", vbYesNo, "VERWIJDEREN PROJECT")
        If antwoord = vbYes Then
            p.synergy = synergy
            p.Vestiging = Vestiging
            p.haalop
            p.verwijderenProject
            MaakArchiefLijst
        End If
    End If
    
End Sub







Private Sub CommandButton5_Click()
Dim wbName As String
Dim r As Long: r = 2

Workbooks.Add
wbName = ActiveWorkbook.Name


Workbooks(wbName).Sheets(1).Range("A1") = "Synergy"
Workbooks(wbName).Sheets(1).Range("B1") = "Vestiging"
Workbooks(wbName).Sheets(1).Range("C1") = "Omschrijving"
Workbooks(wbName).Sheets(1).Range("D1") = "Opdrachtgever"
Workbooks(wbName).Sheets(1).Range("E1") = "PV"
Workbooks(wbName).Sheets(1).Range("F1") = "PL"
Workbooks(wbName).Sheets(1).Range("G1") = "CALC"
Workbooks(wbName).Sheets(1).Range("H1") = "WVB"
Workbooks(wbName).Sheets(1).Range("I1") = "UITV"
Workbooks(wbName).Sheets(1).Range("J1") = "OFFERTE"
Workbooks(wbName).Sheets(1).Range("K1") = "Nabeldatum"

For x = 0 To UBound(lijst, 2)
    Workbooks(wbName).Sheets(1).Range("A" & r) = lijst(0, x)
    Workbooks(wbName).Sheets(1).Range("B" & r) = lijst(10, x)
    Workbooks(wbName).Sheets(1).Range("C" & r) = lijst(1, x)
    Workbooks(wbName).Sheets(1).Range("D" & r) = lijst(2, x)
    Workbooks(wbName).Sheets(1).Range("E" & r) = lijst(3, x)
    Workbooks(wbName).Sheets(1).Range("F" & r) = lijst(4, x)
    Workbooks(wbName).Sheets(1).Range("G" & r) = lijst(5, x)
    Workbooks(wbName).Sheets(1).Range("H" & r) = lijst(6, x)
    Workbooks(wbName).Sheets(1).Range("I" & r) = lijst(7, x)
    Workbooks(wbName).Sheets(1).Range("J" & r) = lijst(9, x)
    Workbooks(wbName).Sheets(1).Range("K" & r) = lijst(13, x)
    r = r + 1
Next x
    Workbooks(wbName).Sheets(1).UsedRange.Columns.AutoFit
    Workbooks(wbName).Sheets(1).UsedRange.AutoFilter
    
End Sub

Private Sub ListBox1_Click()
If ListBox1.ListIndex > -1 Then
    TextBox2 = ListBox1.List(ListBox1.ListIndex, 1)

End If
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim synergy As String

ThisWorkbook.synergy_id = ListBox1.List(ListBox1.ListIndex, 0)
ThisWorkbook.Vestiging = ListBox1.List(ListBox1.ListIndex, 3)
FORM_PROJECT_WIJZIGEN.Show
End Sub

Private Sub TextBox1_Change()
Bijwerkenlijst
End Sub

Private Sub UserForm_Initialize()
MaakArchiefLijst
MaakOverzichtlijst
End Sub

Function MaakArchiefLijst()
Dim db As New DataBase

lijst = db.getLijstBySQL("SELECT * FROM PROJECTEN WHERE STATUS=0 and WACHT=-1 ORDER BY NABELLEN;")
Bijwerkenlijst

End Function

Function MaakOverzichtlijst()
Dim db As New DataBase

overzicht = db.getLijstBySQL("InWachtOverzicht")
bijwerkenoverzicht

End Function

Function UpdateProject()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim db As New DataBase

'Dim l As New Log

   db.Connect
    
    strSQL = "SELECT * FROM PROJECTEN WHERE SYNERGY='" & ListBox1.List(ListBox1.ListIndex, 0) & _
    "' AND Vestiging='" & ListBox1.List(ListBox1.ListIndex, 3) & "';"
    
    rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
        
    If rst.EOF = False And rst.BOF = False Then
        With rst
            .Fields("WACHT").Value = False
            .update
            .Close
        End With
    End If
       
    db.Disconnect
    Set db = Nothing
    'l.createLog "Project uit de wacht gehaald.", pr_updaten, ListBox1.list(ListBox1.ListIndex, 0), project
    
End Function

Function Bijwerkenlijst()
Dim a As Long
Dim x As Long
Dim V As Long: V = 0


ListBox1.Clear

If Not IsEmpty(lijst) = True Then
    If TextBox1 = "" Then
        For x = 0 To UBound(lijst, 2)
            ListBox1.AddItem
            ListBox1.List(x, 0) = lijst(0, x)
            ListBox1.List(x, 1) = lijst(1, x)
            ListBox1.List(x, 2) = lijst(2, x)
            ListBox1.List(x, 3) = lijst(10, x)
            ListBox1.List(x, 4) = lijst(5, x)
            ListBox1.List(x, 5) = lijst(9, x)
            ListBox1.List(x, 6) = Format(lijst(13, x), "dd-mm-yyyy")
            If lijst(13, x) < Now() Then V = V + 1
        Next x
    Else
    a = 0
        For x = 0 To UBound(lijst, 2)
            If InStr(1, lijst(0, x), TextBox1, vbTextCompare) Or InStr(1, lijst(1, x), TextBox1, vbTextCompare) Or InStr(1, lijst(2, x), TextBox1, vbTextCompare) Or InStr(1, lijst(5, x), TextBox1, vbTextCompare) Then
                ListBox1.AddItem
                ListBox1.List(a, 0) = lijst(0, x)
                ListBox1.List(a, 1) = lijst(1, x)
                ListBox1.List(a, 2) = lijst(2, x)
                ListBox1.List(a, 3) = lijst(10, x)
                ListBox1.List(a, 4) = lijst(5, x)
                ListBox1.List(a, 5) = lijst(9, x)
                ListBox1.List(a, 6) = Format(lijst(13, x), "dd-mm-yyyy")
                If lijst(13, x) < Now() Then V = V + 1
                a = a + 1
            End If
        Next x
    End If
End If

LabelAantalInWacht = ListBox1.ListCount
LabelOverTijd = V
End Function

Function bijwerkenoverzicht()
    With ListBox2
    
        .Clear
        For r = 0 To UBound(overzicht, 2)
            .AddItem
            .List(r, 0) = overzicht(1, r)
            .List(r, 1) = overzicht(2, r)
            .List(r, 2) = overzicht(0, r)
        Next r
    End With
End Function
