VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMaterieelorderId As Long
Private pSynergy As String
Private pAanvrager As String
Private pIsvestiging As Boolean
Private pLocatieid As Long
Private pOpmerking As String
Private pStatus As Long
Private pGebruiker As String
Private pStation As String
Private pTijdstip As Date
Private pCorderregels As Collection

'''''''''''''''''''''''''
' MaterieelOrderId property
'''''''''''''''''''''''''

Public Property Get MaterieelOrderId() As Long
    MaterieelOrderId = pMaterieelorderId
End Property

Public Property Let MaterieelOrderId(Value As Long)
    pMaterieelorderId = Value
End Property


'''''''''''''''''''''''''
' Synergy property
'''''''''''''''''''''''''

Public Property Get synergy() As String
    synergy = pSynergy
End Property

Public Property Let synergy(Value As String)
    pSynergy = Value
End Property


'''''''''''''''''''''''''
' Aanvrager property
'''''''''''''''''''''''''

Public Property Get Aanvrager() As String
    Aanvrager = pAanvrager
End Property

Public Property Let Aanvrager(Value As String)
    pAanvrager = Value
End Property


'''''''''''''''''''''''''
' IsVestiging property
'''''''''''''''''''''''''

Public Property Get IsVestiging() As Boolean
    IsVestiging = pIsvestiging
End Property

Public Property Let IsVestiging(Value As Boolean)
    pIsvestiging = Value
End Property


'''''''''''''''''''''''''
' LocatieId property
'''''''''''''''''''''''''

Public Property Get LocatieId() As Long
    LocatieId = pLocatieid
End Property

Public Property Let LocatieId(Value As Long)
    pLocatieid = Value
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


'''''''''''''''''''''''''
' Status property
'''''''''''''''''''''''''

Public Property Get Status() As Long
    Status = pStatus
End Property

Public Property Let Status(Value As Long)
    pStatus = Value
End Property


'''''''''''''''''''''''''
' Gebruiker property
'''''''''''''''''''''''''

Public Property Get Gebruiker() As String
    Gebruiker = pGebruiker
End Property

Public Property Let Gebruiker(Value As String)
    pGebruiker = Value
End Property


'''''''''''''''''''''''''
' Station property
'''''''''''''''''''''''''

Public Property Get Station() As String
    Station = pStation
End Property

Public Property Let Station(Value As String)
    pStation = Value
End Property


'''''''''''''''''''''''''
' Tijdstip property
'''''''''''''''''''''''''

Public Property Get Tijdstip() As Date
    Tijdstip = pTijdstip
End Property

Public Property Let Tijdstip(Value As Date)
    pTijdstip = Value
End Property


'''''''''''''''''''''''''
' cOrderregels property
'''''''''''''''''''''''''

Public Property Get cOrderregels() As Collection
    Set cOrderregels = pCorderregels
End Property

Public Property Let cOrderregels(Value As Collection)
    Set pCorderregels = Value
End Property


'''''''''''''''''''''''''
' cOrderregels
'''''''''''''''''''''''''

Property Get aantalcOrderregels() As Long
    aantalcOrderregels = cOrderregels.Count
End Property

Property Get aantalcOrderregelsGepland() As Long
    Dim mr As New MaterieelOrderRegel
    For Each mr In cOrderregels
        If mr.Gepland = True Then aantalcOrderregelsGepland = aantalcOrderregelsGepland + 1
    Next mr
    
End Property

Property Get itemcOrderregels(ByVal vID As Variant) As MaterieelOrderRegel
    Set itemcOrderregels = cOrderregels(vID)
End Property

Public Sub ToevoegencOrderregels(ByRef item As MaterieelOrderRegel)
    cOrderregels.Add item, CStr(item.MaterieelOrderRegelId) 'eventueel key nog toevoegen.
End Sub

'''''''''''''''''''''''''
' Constructor / Disposal
'''''''''''''''''''''''''

Private Sub class_initialize()
    Set pCorderregels = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCorderregels = Nothing
End Sub


Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New DataBase
Dim maxid As Variant

'Instert Statement
strSQL = "INSERT INTO [MATERIEELORDERS] (" & _
            "[Synergy], " & _
            "[Aanvrager], " & _
            "[IsVestiging], " & _
            "[LocatieId], " & _
            "[Opmerking], " & _
            "[Status], " & _
            "[Gebruiker], " & _
            "[Station], " & _
            "[Tijdstip]) " & _
"VALUES (p2, p3, p4, p5, p6, p7, p8, p9, p10);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Aanvrager) + 1, Me.Aanvrager)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.IsVestiging)), Me.IsVestiging)
    .Parameters.Append .CreateParameter("p5", adInteger, adParamInput, Len(CStr(Me.LocatieId)), Me.LocatieId)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Opmerking) + 1, Me.Opmerking)
    .Parameters.Append .CreateParameter("p7", adInteger, adParamInput, Len(CStr(Me.Status)), Me.Status)
    .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, Len(Me.Gebruiker) + 1, Me.Gebruiker)
    .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, Len(Me.Station) + 1, Me.Station)
    .Parameters.Append .CreateParameter("p10", adDate, adParamInput, Len(CStr(Me.Tijdstip)), Me.Tijdstip)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([MaterieelOrderId]) FROM [MATERIEELORDERS];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.MaterieelOrderId = rs.Fields(0).Value
Set rs = Nothing

'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function

Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
insert = False
Resume Err_Insert

End Function

Public Function update() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New DataBase
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [MATERIEELORDERS] " & _
         "SET " & _
         "[Synergy] = ?, " & _
         "[Aanvrager] = ?, " & _
         "[IsVestiging] = ?, " & _
         "[LocatieId] = ?, " & _
         "[Opmerking] = ?, " & _
         "[Status] = ?, " & _
         "[Gebruiker] = ?, " & _
         "[Station] = ?, " & _
         "[Tijdstip] = ? " & _
         "WHERE [MaterieelOrderId] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Aanvrager) + 1, Me.Aanvrager)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.IsVestiging)), Me.IsVestiging)
    .Parameters.Append .CreateParameter("p5", adInteger, adParamInput, Len(CStr(Me.LocatieId)), Me.LocatieId)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Opmerking) + 1, Me.Opmerking)
    .Parameters.Append .CreateParameter("p7", adInteger, adParamInput, Len(CStr(Me.Status)), Me.Status)
    .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, Len(Me.Gebruiker) + 1, Me.Gebruiker)
    .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, Len(Me.Station) + 1, Me.Station)
    .Parameters.Append .CreateParameter("p10", adDate, adParamInput, Len(CStr(Me.Tijdstip)), Me.Tijdstip)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderId)), Me.MaterieelOrderId)

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
Call errorhandler_MsgBox("Class: " & TypeName(Me) & " , Function: Insert()")
update = False
Resume Err_Insert
End Function

Public Function delete() As Boolean
On Error GoTo ErrorHandler
Dim sql As String
Dim rst As ADODB.Recordset
Dim db As New DataBase

sql = "SELECT * FROM MATERIEELORDERS WHERE MaterieelOrderId = " & Me.MaterieelOrderId & ";"
'We are using a select statement that will return only
'one record (TeacherID 7)

db.Connect

Set rst = New ADODB.Recordset
rst.Open sql, db.connection, adOpenDynamic, adLockOptimistic
'Open RecordSet

With rst

If Not .BOF And Not .EOF Then
    'Ensure that the recordset contains records
    'If no records the code inside the if...end if
    'statement won't run
    
    .MoveLast
    .MoveFirst
    'Not necessary but good practice
    
    If .Supports(adDelete) Then
        'It is possible that the record you want to update
        'is locked by another user. If we don't check before
        'updating, we will generate an error
        
        .delete
        'The only command we need!
        delete = True
    End If
End If

.Close
'Make sure you close the recordset...
End With

db.Disconnect

ExitSub:
Set rst = Nothing
Exit Function
ErrorHandler:
Resume ExitSub

End Function
Function FromList(r As Long, lijst As Variant)
    Me.MaterieelOrderId = lijst(0, r)
    Me.synergy = lijst(1, r)
    Me.Aanvrager = lijst(2, r)
    Me.IsVestiging = lijst(3, r)
    Me.LocatieId = lijst(4, r)
    Me.Opmerking = lijst(5, r)
    Me.Status = lijst(6, r)
    Me.Gebruiker = lijst(7, r)
    Me.Station = lijst(8, r)
    Me.Tijdstip = lijst(9, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELORDER {"
tekst = tekst & vbNewLine & vbTab & "MaterieelOrderId = " & Me.MaterieelOrderId
tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & vbNewLine & vbTab & "Aanvrager = " & Me.Aanvrager
tekst = tekst & vbNewLine & vbTab & "IsVestiging = " & Me.IsVestiging
tekst = tekst & vbNewLine & vbTab & "LocatieId = " & Me.LocatieId
tekst = tekst & vbNewLine & vbTab & "Opmerking = " & Me.Opmerking
tekst = tekst & vbNewLine & vbTab & "Status = " & Me.Status
tekst = tekst & vbNewLine & vbTab & "Gebruiker = " & Me.Gebruiker
tekst = tekst & vbNewLine & vbTab & "Station = " & Me.Station
tekst = tekst & vbNewLine & vbTab & "Tijdstip = " & Me.Tijdstip
tekst = tekst & vbNewLine & vbTab & "cOrderregels = " & Me.aantalcOrderregels
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByMaterieelOrderId()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderId)), Me.MaterieelOrderId)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM MATERIEELORDERS WHERE MaterieelOrderId = p1;", pr)

    FromList 0, lijst

End Function

