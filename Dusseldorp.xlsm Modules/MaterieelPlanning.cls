VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId As Long
Private pMaterieelid As Long
Private pStartDatum As Date
Private pEindDatum As Date
Private pMaterieelsoortId As Long
Private pGekoppeld As Boolean
Private pSynergy As String
Private pMaterieel As New Materieel
Private pMaterieelsoort As New MaterieelSoort
Private pIsGepickt As Boolean
Private pKolomnummerStart As Long
Private pKolomnummerEind As Long


'''''''''''''''''''''''''
' Id property
'''''''''''''''''''''''''

Public Property Get Id() As Long
    Id = pId
End Property

Public Property Let Id(Value As Long)
    pId = Value
End Property


'''''''''''''''''''''''''
' MaterieelId property
'''''''''''''''''''''''''

Public Property Get MaterieelId() As Long
    MaterieelId = pMaterieelid
End Property

Public Property Let MaterieelId(Value As Long)
    pMaterieelid = Value
End Property


'''''''''''''''''''''''''
' StartDatum property
'''''''''''''''''''''''''

Public Property Get startdatum() As Date
    startdatum = pStartDatum
End Property

Public Property Let startdatum(Value As Date)
    pStartDatum = Value
End Property

'''''''''''''''''''''''''
' EindDatum property
'''''''''''''''''''''''''

Public Property Get einddatum() As Date
    einddatum = pEindDatum
End Property

Public Property Let einddatum(Value As Date)
    pEindDatum = Value
End Property


'''''''''''''''''''''''''
' MaterieelSoort property
'''''''''''''''''''''''''

Public Property Get MaterieelSoortId() As Long
    MaterieelSoortId = pMaterieelsoortId
End Property

Public Property Let MaterieelSoortId(Value As Long)
    pMaterieelsoortId = Value
End Property


'''''''''''''''''''''''''
' Gekoppeld property
'''''''''''''''''''''''''

Public Property Get Gekoppeld() As Boolean
    Gekoppeld = pGekoppeld
End Property

Public Property Let Gekoppeld(Value As Boolean)
    pGekoppeld = Value
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
' Materieel property
'''''''''''''''''''''''''

Public Property Get Materieel() As Materieel
   Set Materieel = pMaterieel
End Property

Public Property Let Materieel(Value As Materieel)
   Set pMaterieel = Value
End Property

'''''''''''''''''''''''''
' MaterieelSoort property
'''''''''''''''''''''''''

Public Property Get MaterieelSoort() As MaterieelSoort
   Set MaterieelSoort = pMaterieelsoort
End Property

Public Property Let MaterieelSoort(Value As MaterieelSoort)
   Set pMaterieelsoort = Value
End Property

'''''''''''''''''''''''''
' Kolomnummer property
'''''''''''''''''''''''''

Public Property Get KolomnummerStart() As Long
    KolomnummerStart = pKolomnummerStart
End Property

Public Property Let KolomnummerStart(Value As Long)
    pKolomnummerStart = Value
End Property


'''''''''''''''''''''''''
' Kolomnummer property
'''''''''''''''''''''''''

Public Property Get KolomnummerEind() As Long
    KolomnummerEind = pKolomnummerEind
End Property

Public Property Let KolomnummerEind(Value As Long)
    pKolomnummerEind = Value
End Property

'''''''''''''''''''''''''
' isGepicket property
'''''''''''''''''''''''''

Public Property Get isGepickt() As Boolean
    isGepickt = pIsGepickt
End Property

Public Property Let isGepickt(Value As Boolean)
    pIsGepickt = Value
End Property


Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New DataBase
Dim maxid As Variant

'Instert Statement
strSQL = "INSERT INTO [PLANNING_MATERIEEL] (" & _
            "[MaterieelId], " & _
            "[StartDatum], " & _
            "[EindDatum], " & _
            "[MaterieelSoortId], " & _
            "[Gekoppeld], " & _
            "[Synergy]) " & _
"VALUES (p2, p3, p4, p5, p6, p7);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText

    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
    .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
    .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
    .Parameters.Append .CreateParameter("p5", adInteger, adParamInput, Len(CStr(Me.MaterieelSoortId)), Me.MaterieelSoortId)
    .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.Gekoppeld)), Me.Gekoppeld)
    .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([Id]) FROM [PLANNING_MATERIEEL];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.Id = rs.Fields(0).Value
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
strSQL = "UPDATE [PLANNING_MATERIEEL] " & _
         "SET " & _
         "[MaterieelId] = ?, " & _
         "[StartDatum] = ?, " & _
         "[EindDatum] = ?, " & _
         "[MaterieelSoortId] = ?, " & _
         "[Gekoppeld] = ?, " & _
         "[Synergy] = ?, " & _
         "[IsGepickt] = ? " & _
         "WHERE [Id] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
    .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
    .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
    .Parameters.Append .CreateParameter("p5", adInteger, adParamInput, Len(CStr(Me.MaterieelSoortId)), Me.MaterieelSoortId)
    .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.Gekoppeld)), Me.Gekoppeld)
    .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.isGepickt)), Me.isGepickt)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.Id)), Me.Id)

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

sql = "SELECT * FROM PLANNING_MATERIEEL WHERE Id = " & Me.Id & ";"
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
    Me.Id = lijst(0, r)
    Me.MaterieelId = lijst(1, r)
    Me.startdatum = lijst(2, r)
    Me.einddatum = lijst(3, r)
    Me.MaterieelSoortId = lijst(4, r)
    Me.Gekoppeld = lijst(5, r)
    Me.synergy = lijst(6, r)
    Me.isGepickt = lijst(7, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELPLANNING {"
tekst = tekst & vbNewLine & vbTab & "Id = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "MaterieelId = " & Me.MaterieelId
tekst = tekst & vbNewLine & vbTab & "StartDatum = " & Me.startdatum
tekst = tekst & vbNewLine & vbTab & "EindDatum = " & Me.einddatum
tekst = tekst & vbNewLine & vbTab & "MaterieelSoorti = " & Me.MaterieelSoortId
tekst = tekst & vbNewLine & vbTab & "Gekoppeld = " & Me.Gekoppeld
tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & vbNewLine & vbTab & "IsGepickt = " & Me.isGepickt
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetById()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.Id)), Me.Id)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM PLANNING_MATERIEEL WHERE Id = ?;", pr)

    FromList 0, lijst

End Function

Public Function DeleteDatumMaterieel() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
    db.Connect
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_MATERIEEL] " & _
             "WHERE [MaterieelId] = ? AND [StartDatum] = ? AND [EindDatum] = ?;"
    
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
        .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    db.Disconnect
    
    If lRecordsAffected = 0 Then
      DeleteDatumMaterieel = False
    Else
      'Return a good return code
      DeleteDatumMaterieel = True
    End If
    
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumMaterieel = False
    Resume Exit_Insert

End Function


Public Function DeleteDatumMaterieelSynergy() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
    db.Connect
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_MATERIEEL] " & _
             "WHERE [MaterieelId] = ? AND [Datum] = ? AND [Datum] = ? AND [Synergy] = ? AND [MaterieelSoortId] = ?;"
    
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
        .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p5", adBigInt, adParamInput, Len(CStr(Me.MaterieelSoortId)), Me.MaterieelSoortId)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    db.Disconnect
    
    If lRecordsAffected = 0 Then
      DeleteDatumMaterieelSynergy = False
    Else
      'Return a good return code
      DeleteDatumMaterieelSynergy = True
    End If
    
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumMaterieelSynergy = False
    Resume Exit_Insert
End Function



Public Function DeleteDatumMaterieelMaterieelsoort() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
   db.Connect
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_MATERIEEL] " & _
             "WHERE [MaterieelId] = ? AND [Datum] = ? AND [MaterieelSoortId] = ?;"
    
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
        .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
        .Parameters.Append .CreateParameter("p4", adBigInt, adParamInput, Len(CStr(Me.MaterieelSoortId)), Me.MaterieelSoortId)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      DeleteDatumMaterieelMaterieelsoort = False
    Else
      'Return a good return code
      DeleteDatumMaterieelMaterieelsoort = True
    End If
    
    db.Disconnect
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumMaterieelMaterieelsoort = False
    Resume Exit_Insert
End Function
