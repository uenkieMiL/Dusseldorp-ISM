VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_BESTELLINGEN 
   Caption         =   "UserForm1"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "FORM_BESTELLINGEN.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_BESTELLINGEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lijstOrders As New Collection
Public gekozenOrder As MaterieelOrder
Public CKalender As New Collection

Private Sub ListBoxOrders_Click()
If Functies.InCollection(lijstOrders, ListBoxOrders.List(ListBoxOrders.ListIndex, 0)) = True Then
    Set gekozenOrder = lijstOrders.item(CStr(ListBoxOrders.List(ListBoxOrders.ListIndex, 0)))
    HaalLijstOp
    Bijwerkenlijst
Else
    Set gekozenOrder = Nothing
End If
End Sub

Private Sub ListBoxOrders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    FORM_BESTELLING_INZIEN.Show
End Sub

Private Sub UserForm_Initialize()
    Dim vandaag As Date, mindatum As Date, maxdatum As Date
    vandaag = Now()
    mindatum = DateAdd("d", 0 - Weekday(vandaag, vbMonday) - 14, Now())
    maxdatum = DateAdd("d", (104 * 7 - 1), mindatum)
    Set CKalender = Lijsten.KalenderStartEind(mindatum, maxdatum)
    HaalLijstOp
    Bijwerkenlijst
End Sub

Public Function HaalLijstOp()
     Set lijstOrders = Lijsten.MaakLijstInTePlannenMaterieel
End Function

Public Function Bijwerkenlijst()
Dim o As MaterieelOrder

    ListBoxOrders.Clear
    For Each o In lijstOrders
        With ListBoxOrders
        .AddItem
        .List(.ListCount - 1, 0) = o.MaterieelOrderId
        .List(.ListCount - 1, 1) = o.synergy
        .List(.ListCount - 1, 2) = o.Aanvrager
        If o.IsVestiging = True Then .List(.ListCount - 1, 3) = "x"
        .List(.ListCount - 1, 4) = o.Tijdstip
        .List(.ListCount - 1, 5) = o.aantalcOrderregelsGepland & "/" & o.aantalcOrderregels
        .List(.ListCount - 1, 6) = o.Opmerking
        End With
    Next o
End Function
