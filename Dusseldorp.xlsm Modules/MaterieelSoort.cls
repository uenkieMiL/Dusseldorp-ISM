VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelSoort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId As Long
Private pOmschrijving As String
Private pKleur As Long
Private pKoppelbaar As Boolean
Private pInplanbaar As Boolean

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
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property


'''''''''''''''''''''''''
' Kleur property
'''''''''''''''''''''''''

Public Property Get Kleur() As Long
    Kleur = pKleur
End Property

Public Property Let Kleur(Value As Long)
    pKleur = Value
End Property


'''''''''''''''''''''''''
' Koppelbaar property
'''''''''''''''''''''''''

Public Property Get Koppelbaar() As Boolean
    Koppelbaar = pKoppelbaar
End Property

Public Property Let Koppelbaar(Value As Boolean)
    pKoppelbaar = Value
End Property


'''''''''''''''''''''''''
' Inplanbaar property
'''''''''''''''''''''''''

Public Property Get Inplanbaar() As Boolean
    Inplanbaar = pInplanbaar
End Property

Public Property Let Inplanbaar(Value As Boolean)
    pInplanbaar = Value
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
strSQL = "INSERT INTO [MATERIEELSOORT] (" & _
            "[Id], " & _
            "[Omschrijving], " & _
            "[Kleur], " & _
            "[Koppelbaar], " & _
            "[Inplanbaar]) " & _
"VALUES (p1, p2, p3, p4, p5);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText

    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.Id)), Me.Id)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.Kleur)), Me.Kleur)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Koppelbaar)), Me.Koppelbaar)
    .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Inplanbaar)), Me.Inplanbaar)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([Id]) FROM [MATERIEELSOORT];"
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
strSQL = "UPDATE [MATERIEELSOORT] " & _
         "SET " & _
         "[Omschrijving] = ?, " & _
         "[Kleur] = ?, " & _
         "[Koppelbaar] = ?, " & _
         "[Inplanbaar] = ? " & _
         "WHERE [Id] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.Kleur)), Me.Kleur)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Koppelbaar)), Me.Koppelbaar)
    .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Inplanbaar)), Me.Inplanbaar)
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

sql = "SELECT * FROM MATERIEELSOORT WHERE Id = " & Me.Id & ";"
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
    Me.Omschrijving = lijst(1, r)
    Me.Kleur = lijst(2, r)
    Me.Koppelbaar = lijst(3, r)
    Me.Inplanbaar = lijst(4, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELSOORT {"
tekst = tekst & vbNewLine & vbTab & "Id = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Kleur = " & Me.Kleur
tekst = tekst & vbNewLine & vbTab & "Koppelbaar = " & Me.Koppelbaar
tekst = tekst & vbNewLine & vbTab & "Inplanbaar = " & Me.Inplanbaar
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

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM MATERIEELSOORT WHERE Id = p1;", pr)

    FromList 0, lijst

End Function

