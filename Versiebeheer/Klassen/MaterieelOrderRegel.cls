VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelOrderRegel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMaterieelorderregelId As Long
Private pMaterieelorderId As Long
Private pMaterieeltypeid As Long
Private pStartDatum As Date
Private pEindDatum As Date
Private pGepland As Boolean
Private pGeplanddatum As Date
Private pGeplandgebruiker As String
Private pGeplandstation As String
Private pPlanning_MaterieelId As Long
Private pMaterieelType As New MaterieelType

'''''''''''''''''''''''''
' MaterieelOrderRegelId property
'''''''''''''''''''''''''

Public Property Get MaterieelOrderRegelId() As Long
    MaterieelOrderRegelId = pMaterieelorderregelId
End Property

Public Property Let MaterieelOrderRegelId(Value As Long)
    pMaterieelorderregelId = Value
End Property


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
' MaterieelTypeId property
'''''''''''''''''''''''''

Public Property Get MaterieelTypeId() As Long
    MaterieelTypeId = pMaterieeltypeid
End Property

Public Property Let MaterieelTypeId(Value As Long)
    pMaterieeltypeid = Value
End Property


'''''''''''''''''''''''''
' Startdatum property
'''''''''''''''''''''''''

Public Property Get startdatum() As Date
    startdatum = pStartDatum
End Property

Public Property Let startdatum(Value As Date)
    pStartDatum = Value
End Property


'''''''''''''''''''''''''
' Einddatum property
'''''''''''''''''''''''''

Public Property Get einddatum() As Date
    einddatum = pEindDatum
End Property

Public Property Let einddatum(Value As Date)
    pEindDatum = Value
End Property


'''''''''''''''''''''''''
' Gepland property
'''''''''''''''''''''''''

Public Property Get Gepland() As Boolean
    Gepland = pGepland
End Property

Public Property Let Gepland(Value As Boolean)
    pGepland = Value
End Property


'''''''''''''''''''''''''
' GeplandDatum property
'''''''''''''''''''''''''

Public Property Get GeplandDatum() As Date
    GeplandDatum = pGeplanddatum
End Property

Public Property Let GeplandDatum(Value As Date)
    pGeplanddatum = Value
End Property


'''''''''''''''''''''''''
' GeplandGebruiker property
'''''''''''''''''''''''''

Public Property Get GeplandGebruiker() As String
    GeplandGebruiker = pGeplandgebruiker
End Property

Public Property Let GeplandGebruiker(Value As String)
    pGeplandgebruiker = Value
End Property


'''''''''''''''''''''''''
' GeplandStation property
'''''''''''''''''''''''''

Public Property Get GeplandStation() As String
    GeplandStation = pGeplandstation
End Property

Public Property Let GeplandStation(Value As String)
    pGeplandstation = Value
End Property

'''''''''''''''''''''''''
' MaterieelType property
'''''''''''''''''''''''''

Public Property Get MaterieelType() As MaterieelType
    Set MaterieelType = pMaterieelType
End Property

Public Property Let MaterieelType(Value As MaterieelType)
    pMaterieelType = Value
End Property


'''''''''''''''''''''''''
' Planning_MaterieelId property
'''''''''''''''''''''''''

Public Property Get Planning_MaterieelId() As Long
    Planning_MaterieelId = pPlanning_MaterieelId
End Property

Public Property Let Planning_MaterieelId(Value As Long)
    pPlanning_MaterieelId = Value
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
strSQL = "INSERT INTO [MATERIEELORDERREGELS] (" & _
            "[MaterieelOrderId], " & _
            "[MaterieelTypeId], " & _
            "[Startdatum], " & _
            "[Einddatum]) " & _
"VALUES (p2, p3, p5, p6);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderId)), Me.MaterieelOrderId)
    .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.MaterieelTypeId)), Me.MaterieelTypeId)
    .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
    .Parameters.Append .CreateParameter("p6", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([MaterieelOrderRegelId]) FROM [MATERIEELORDERREGELS];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.MaterieelOrderRegelId = rs.Fields(0).Value
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
strSQL = "UPDATE [MATERIEELORDERREGELS] " & _
         "SET " & _
         "[MaterieelOrderId] = ?, " & _
         "[MaterieelTypeId] = ?, " & _
         "[Startdatum] = ?, " & _
         "[Einddatum] = ?, " & _
         "[Gepland] = ?, " & _
         "[GeplandDatum] = ?, " & _
         "[GeplandGebruiker] = ?, " & _
         "[GeplandStation] = ?, " & _
         "[Planning_MaterieelId] = ? " & _
         "WHERE [MaterieelOrderRegelId] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderId)), Me.MaterieelOrderId)
    .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.MaterieelTypeId)), Me.MaterieelTypeId)
    .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
    .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
    .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.Gepland)), Me.Gepland)
    .Parameters.Append .CreateParameter("p7", adDate, adParamInput, Len(CStr(Me.GeplandDatum)), Me.GeplandDatum)
    .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, Len(Me.GeplandGebruiker) + 1, Me.GeplandGebruiker)
    .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, Len(Me.GeplandStation) + 1, Me.GeplandStation)
    .Parameters.Append .CreateParameter("p10", adInteger, adParamInput, Len(CStr(Me.Planning_MaterieelId)), Me.Planning_MaterieelId)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderRegelId)), Me.MaterieelOrderRegelId)

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

sql = "SELECT * FROM MATERIEELORDERREGELS WHERE MaterieelOrderRegelId = " & Me.MaterieelOrderRegelId & ";"
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

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELORDERREGEL {"
tekst = tekst & vbNewLine & vbTab & "MaterieelOrderRegelId = " & Me.MaterieelOrderRegelId
tekst = tekst & vbNewLine & vbTab & "MaterieelOrderId = " & Me.MaterieelOrderId
tekst = tekst & vbNewLine & vbTab & "MaterieelTypeId = " & Me.MaterieelTypeId
tekst = tekst & vbNewLine & vbTab & "Startdatum = " & Me.startdatum
tekst = tekst & vbNewLine & vbTab & "Einddatum = " & Me.einddatum
tekst = tekst & vbNewLine & vbTab & "Gepland = " & Me.Gepland
tekst = tekst & vbNewLine & vbTab & "GeplandDatum = " & Me.GeplandDatum
tekst = tekst & vbNewLine & vbTab & "GeplandGebruiker = " & Me.GeplandGebruiker
tekst = tekst & vbNewLine & vbTab & "GeplandStation = " & Me.GeplandStation
tekst = tekst & vbNewLine & vbTab & "Aantal = " & Me.Planning_MaterieelId
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByMaterieelOrderRegelId()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelOrderRegelId)), Me.MaterieelOrderRegelId)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM MATERIEELORDERREGELS WHERE MaterieelOrderRegelId = p1;", pr)

    FromList 0, lijst

End Function


Function FromListExtra(r As Long, lijst As Variant, Optional v As Long = 0)
    Me.MaterieelOrderRegelId = lijst(0 + v, r)
    Me.MaterieelOrderId = lijst(1 + v, r)
    Me.MaterieelTypeId = lijst(2 + v, r)
    Me.startdatum = lijst(3 + v, r)
    Me.einddatum = lijst(4 + v, r)
    Me.Gepland = lijst(5 + v, r)
    If IsNull(lijst(6 + v, r)) = False Then Me.GeplandDatum = lijst(6 + v, r)
    If IsNull(lijst(7 + v, r)) = False Then Me.GeplandGebruiker = lijst(7 + v, r)
    If IsNull(lijst(8 + v, r)) = False Then Me.GeplandStation = lijst(8 + v, r)
    If IsNull(lijst(9 + v, r)) = False Then Me.GeplandStation = lijst(9 + v, r)
End Function

