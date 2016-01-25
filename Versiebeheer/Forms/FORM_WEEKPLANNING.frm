VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORM_WEEKPLANNING 
   Caption         =   "MAAK WEEK PLANNING"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3540
   OleObjectBlob   =   "FORM_WEEKPLANNING.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORM_WEEKPLANNING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ToggleButton1_Click()
If CheckBox1 = True Then ProjectPlanning.acquesitie = True
If CheckBox2 = True Then ProjectPlanning.calculatie = True
If CheckBox4 = True Then ProjectPlanning.uitvoering = True
Unload Me
End Sub

Private Sub UserForm_Initialize()
TextBox1 = ProjectPlanning.form_synergy_voorbereidingsplanning
End Sub
