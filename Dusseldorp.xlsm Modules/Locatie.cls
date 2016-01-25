VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Locatie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pLocatieid As Long
Private pSynergy As String
Private pOmschrijivng As String
Private pAdres As String
Private pPostcode As String
Private pPlaats As String




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
' Synergy property
'''''''''''''''''''''''''

Public Property Get synergy() As String
    synergy = pSynergy
End Property

Public Property Let synergy(Value As String)
    pSynergy = Value
End Property


'''''''''''''''''''''''''
' Omschrijivng property
'''''''''''''''''''''''''

Public Property Get Omschrijivng() As String
    Omschrijivng = pOmschrijivng
End Property

Public Property Let Omschrijivng(Value As String)
    pOmschrijivng = Value
End Property


'''''''''''''''''''''''''
' Adres property
'''''''''''''''''''''''''

Public Property Get Adres() As String
    Adres = pAdres
End Property

Public Property Let Adres(Value As String)
    pAdres = Value
End Property


'''''''''''''''''''''''''
' Postcode property
'''''''''''''''''''''''''

Public Property Get Postcode() As String
    Postcode = pPostcode
End Property

Public Property Let Postcode(Value As String)
    pPostcode = Value
End Property


'''''''''''''''''''''''''
' Plaats property
'''''''''''''''''''''''''

Public Property Get Plaats() As String
    Plaats = pPlaats
End Property

Public Property Let Plaats(Value As String)
    pPlaats = Value
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
strSQL = "INSERT INTO [LOCATIES] (" & _
            "[Synergy], " & _
            "[Omschrijivng], " & _
            "[Adres], " & _
            "[Postcode], " & _
            "[Plaats]) " & _
"VALUES (p2, p3, p4, p5, p6);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijivng) + 1, Me.Omschrijivng)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Adres) + 1, Me.Adres)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Postcode) + 1, Me.Postcode)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Plaats) + 1, Me.Plaats)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([LocatieId]) FROM [LOCATIES];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.LocatieId = rs.Fields(0).Value
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
strSQL = "UPDATE [LOCATIES] " & _
         "SET " & _
         "[Synergy] = ?, " & _
         "[Omschrijivng] = ?, " & _
         "[Adres] = ?, " & _
         "[Postcode] = ?, " & _
         "[Plaats] = ? " & _
         "WHERE [LocatieId] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijivng) + 1, Me.Omschrijivng)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Adres) + 1, Me.Adres)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Postcode) + 1, Me.Postcode)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Plaats) + 1, Me.Plaats)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.LocatieId)), Me.LocatieId)

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

sql = "SELECT * FROM LOCATIES WHERE LocatieId = " & Me.LocatieId & ";"
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
    Me.LocatieId = lijst(0, r)
    Me.synergy = lijst(1, r)
    Me.Omschrijivng = lijst(2, r)
    Me.Adres = lijst(3, r)
    Me.Postcode = lijst(4, r)
    Me.Plaats = lijst(5, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "LOACTIE {"
tekst = tekst & vbNewLine & vbTab & "LocatieId = " & Me.LocatieId
tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & vbNewLine & vbTab & "Omschrijivng = " & Me.Omschrijivng
tekst = tekst & vbNewLine & vbTab & "Adres = " & Me.Adres
tekst = tekst & vbNewLine & vbTab & "Postcode = " & Me.Postcode
tekst = tekst & vbNewLine & vbTab & "Plaats = " & Me.Plaats
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByLocatieId()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.LocatieId)), Me.LocatieId)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM LOCATIES WHERE LocatieId = p1;", pr)

    FromList 0, lijst

End Function

