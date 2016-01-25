VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KantoorPersoneel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pAfkorting As String
Private Pnaam As String
Private pGebruikersnaam As String
Private pEmail As String
Private pPV As Boolean
Private pPL As Boolean
Private pCALC As Boolean
Private pWVB As Boolean
Private pUITV As Boolean
Private pNAB As Boolean
Private pOFFERTE As Boolean
Private pInactief As Boolean
Private pOpmerking As String

'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get afkorting() As String
    afkorting = pAfkorting
End Property

Public Property Let afkorting(Value As String)
    pAfkorting = Value
End Property

'''''''''''''''''''''''''
' Naam property
'''''''''''''''''''''''''

Public Property Get Naam() As String
    Naam = Pnaam
End Property

Public Property Let Naam(Value As String)
    Pnaam = Value
End Property

'''''''''''''''''''''''''
' Gebruikersnaam property
'''''''''''''''''''''''''

Public Property Get Gebruikersnaam() As String
    Gebruikersnaam = pGebruikersnaam
End Property

Public Property Let Gebruikersnaam(Value As String)
    pGebruikersnaam = Value
End Property

'''''''''''''''''''''''''
' Email property
'''''''''''''''''''''''''

Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(Value As String)
    pEmail = Value
End Property

'''''''''''''''''''''''''
' PV property
'''''''''''''''''''''''''

Public Property Get pv() As Boolean
    pv = pPV
End Property

Public Property Let pv(Value As Boolean)
    pPV = Value
End Property

'''''''''''''''''''''''''
' PL property
'''''''''''''''''''''''''

Public Property Get pl() As Boolean
    pl = pPL
End Property

Public Property Let pl(Value As Boolean)
    pPL = Value
End Property

'''''''''''''''''''''''''
' CALC property
'''''''''''''''''''''''''

Public Property Get CALC() As Boolean
    CALC = pCALC
End Property

Public Property Let CALC(Value As Boolean)
    pCALC = Value
End Property

'''''''''''''''''''''''''
' WVB property
'''''''''''''''''''''''''

Public Property Get wvb() As Boolean
    wvb = pWVB
End Property

Public Property Let wvb(Value As Boolean)
    pWVB = Value
End Property

'''''''''''''''''''''''''
' UITV property
'''''''''''''''''''''''''

Public Property Get uitv() As Boolean
    uitv = pUITV
End Property

Public Property Let uitv(Value As Boolean)
    pUITV = Value
End Property

'''''''''''''''''''''''''
' NAB property
'''''''''''''''''''''''''

Public Property Get NAB() As Boolean
    NAB = pNAB
End Property

Public Property Let NAB(Value As Boolean)
    pNAB = Value
End Property

'''''''''''''''''''''''''
' OFFERTE property
'''''''''''''''''''''''''

Public Property Get OFFERTE() As Boolean
    OFFERTE = pOFFERTE
End Property

Public Property Let OFFERTE(Value As Boolean)
    pOFFERTE = Value
End Property

'''''''''''''''''''''''''
' INACTIEF property
'''''''''''''''''''''''''

Public Property Get Inactief() As Boolean
    Inactief = pInactief
End Property

Public Property Let Inactief(Value As Boolean)
    pInactief = Value
End Property

'''''''''''''''''''''''''
' Opmerking property
'''''''''''''''''''''''''

Public Property Get Opmerking() As String
    Opmerking = pOpmerking
End Property
Public Property Let Opmerking(Value As String)
    pOpmerking = Value
End Property



Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New DataBase
Dim maxid As Variant

'Instert Statement
strSQL = "INSERT INTO [KantoorPersoneel] (" & _
            "[Afkorting], " & _
            "[Naam], " & _
            "[Gebruikersnaam], " & _
            "[Email], " & _
            "[PV], " & _
            "[PL], " & _
            "[CALC], " & _
            "[WVB], " & _
            "[UITV], " & _
            "[NAB], " & _
            "[OFFERTE], " & _
            "[INACTIEF], " & _
            "[OPMERKING]) " & _
"VALUES (p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.afkorting), Me.afkorting)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Naam) + 1, Me.Naam)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Gebruikersnaam) + 1, Me.Gebruikersnaam)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Email) + 1, Me.Email)
    .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.pv)), Me.pv)
    .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.pl)), Me.pl)
    .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.CALC)), Me.CALC)
    .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.wvb)), Me.wvb)
    .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.uitv)), Me.uitv)
    .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.NAB)), Me.NAB)
    .Parameters.Append .CreateParameter("p11", adBoolean, adParamInput, Len(CStr(Me.OFFERTE)), Me.OFFERTE)
    .Parameters.Append .CreateParameter("p12", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
    .Parameters.Append .CreateParameter("p13", adVarChar, adParamInput, Len(Me.Opmerking) + 1, Me.Opmerking)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True


'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function

Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
insert = False


End Function

Public Function update() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New DataBase
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [KantoorPersoneel] " & _
         "SET " & _
         "[Naam] = ?, " & _
         "[Gebruikersnaam] = ?, " & _
         "[Email] = ?, " & _
         "[PV] = ?, " & _
         "[PL] = ?, " & _
         "[CALC] = ?, " & _
         "[WVB] = ?, " & _
         "[UITV] = ?, " & _
         "[NAB] = ?, " & _
         "[OFFERTE] = ?, " & _
         "[INACTIEF] = ?, " & _
         "[OPMERKING] = ? " & _
         "WHERE [AFKORTING] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL
    .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Naam) + 1, Me.Naam)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Gebruikersnaam) + 1, Me.Gebruikersnaam)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Email) + 1, Me.Email)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.pv)), Me.pv)
    .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.pl)), Me.pl)
    .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.CALC)), Me.CALC)
    .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.wvb)), Me.wvb)
    .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.uitv)), Me.uitv)
    .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.NAB)), Me.NAB)
    .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.OFFERTE)), Me.OFFERTE)
    .Parameters.Append .CreateParameter("p11", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
    .Parameters.Append .CreateParameter("p12", adVarChar, adParamInput, Len(Me.Opmerking) + 1, Me.Opmerking)
    .Parameters.Append .CreateParameter("p13", adVarChar, adParamInput, Len(Me.afkorting), Me.afkorting)

Set rst = .Execute(lRecordsAffected)
End With
    
If lRecordsAffected = 0 Then
    update = False
Else
    'Return a good return code
    update = True
End If

Exit_Insert:
'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function
    
Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
update = False
Resume Exit_Insert
End Function

Public Function HaalOpMetAfkorting()
    
Dim db As New DataBase
Dim rst As New ADODB.Recordset
Dim strSQL As String

strSQL = "SELECT * FROM KantoorPersoneel where Afkorting = '" & Me.afkorting & "';"

db.Connect

Set rst = db.OpenRecordset(strSQL)

Me.afkorting = rst.Fields(0)
Me.Naam = rst.Fields(1)
If IsNull(rst.Fields(2)) = False Then Me.Gebruikersnaam = rst.Fields(2)
If IsNull(rst.Fields(3)) = False Then Me.Email = rst.Fields(3)
Me.pv = rst.Fields(4)
Me.pl = rst.Fields(5)
Me.CALC = rst.Fields(6)
Me.wvb = rst.Fields(7)
Me.uitv = rst.Fields(8)
Me.NAB = rst.Fields(8)
Me.OFFERTE = rst.Fields(9)
Me.Inactief = rst.Fields(10)
Me.Opmerking = rst.Fields(11)
db.Disconnect

Set rst = Nothing
Set db = Nothing

End Function

Function ToString() As String
Dim tekst As String

    tekst = "TEKST {"
    tekst = tekst & vbNewLine & vbTab & "Afkorting = " & Me.afkorting
    tekst = tekst & vbNewLine & vbTab & "Naam = " & Me.Naam
    tekst = tekst & vbNewLine & vbTab & "Gebruikersnaam = " & Me.Gebruikersnaam
    tekst = tekst & vbNewLine & vbTab & "Email = " & Me.Email
    tekst = tekst & vbNewLine & vbTab & "PV = " & Me.pv
    tekst = tekst & vbNewLine & vbTab & "PL = " & Me.pl
    tekst = tekst & vbNewLine & vbTab & "CALC = " & Me.CALC
    tekst = tekst & vbNewLine & vbTab & "WVB = " & Me.wvb
    tekst = tekst & vbNewLine & vbTab & "UITV = " & Me.uitv
    tekst = tekst & vbNewLine & vbTab & "NAB = " & Me.NAB
    tekst = tekst & vbNewLine & vbTab & "OFFERTE = " & Me.OFFERTE
    tekst = tekst & vbNewLine & vbTab & "INACTIEF = " & Me.Inactief
    tekst = tekst & vbNewLine & vbTab & "Opmerking = " & Me.Opmerking
    tekst = tekst & vbNewLine & "}"
    ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function

Function FromList(r As Long, lijst As Variant)
    Me.afkorting = lijst(0, r)
    Me.Naam = lijst(1, r)
    If IsNull(lijst(2, r)) = False Then Me.Gebruikersnaam = lijst(2, r)
    If IsNull(lijst(3, r)) = False Then Me.Email = lijst(3, r)
    Me.pv = lijst(4, r)
    Me.pl = lijst(5, r)
    Me.CALC = lijst(6, r)
    Me.wvb = lijst(7, r)
    Me.uitv = lijst(8, r)
    Me.NAB = lijst(9, r)
    Me.OFFERTE = lijst(10, r)
    Me.Inactief = lijst(11, r)
    If IsNull(lijst(12, r)) = False Then Me.Opmerking = lijst(12, r)
End Function



