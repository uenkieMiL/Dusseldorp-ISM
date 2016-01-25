VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelActie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pActieid As Long
Private pMaterieelid As Long
Private pGebruiker As String
Private pStation As String
Private pTijdstip As Date
Private pOmschrijving As String
Private pWaarde As String


'''''''''''''''''''''''''
' ActieId property
'''''''''''''''''''''''''

Public Property Get ActieId() As Long
    ActieId = pActieid
End Property

Public Property Let ActieId(Value As Long)
    pActieid = Value
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
' Gebruiker property
'''''''''''''''''''''''''

Public Property Get Gebruiker() As String
    Gebruiker = pGebruiker
End Property

Public Property Let Gebruiker(Value As String)
    If pGebruiker = "" Then pGebruiker = Environ$("username")
    pGebruiker = Value
End Property


'''''''''''''''''''''''''
' Station property
'''''''''''''''''''''''''

Public Property Get Station() As String

    Station = pStation
End Property

Public Property Let Station(Value As String)
    If pStation = "" Then pStation = Environ$("computername")
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
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property


'''''''''''''''''''''''''
' Waarde property
'''''''''''''''''''''''''

Public Property Get Waarde() As String
    Waarde = pWaarde
End Property

Public Property Let Waarde(Value As String)
    pWaarde = Value
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
strSQL = "INSERT INTO [MATERIEELACTIE] (" & _
            "[MaterieelId], " & _
            "[Gebruiker], " & _
            "[Station], " & _
            "[Tijdstip], " & _
            "[Omschrijving], " & _
            "[Waarde]) " & _
"VALUES (p1, p2, p3, p4, p5, p6);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Gebruiker) + 1, Me.Gebruiker)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Station) + 1, Me.Station)
    .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.Tijdstip)), Me.Tijdstip)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Waarde) + 1, Me.Waarde)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([ActieId]) FROM [MATERIEELACTIE];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.ActieId = rs.Fields(0).Value
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
strSQL = "UPDATE [MATERIEELACTIE] " & _
         "SET " & _
         "[MaterieelId] = ?, " & _
         "[Gebruiker] = ?, " & _
         "[Station] = ?, " & _
         "[Tijdstip] = ?, " & _
         "[Omschrijving] = ?, " & _
         "[Waarde] = ? " & _
         "WHERE [ActieId] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.MaterieelId)), Me.MaterieelId)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Gebruiker) + 1, Me.Gebruiker)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Station) + 1, Me.Station)
    .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.Tijdstip)), Me.Tijdstip)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, Len(Me.Waarde) + 1, Me.Waarde)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.ActieId)), Me.ActieId)

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

sql = "SELECT * FROM MATERIEELACTIE WHERE ActieId = " & Me.ActieId & ";"
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
    Me.ActieId = lijst(0, r)
    Me.MaterieelId = lijst(1, r)
    Me.Gebruiker = lijst(2, r)
    Me.Station = lijst(3, r)
    Me.Tijdstip = lijst(4, r)
    Me.Omschrijving = lijst(5, r)
    Me.Waarde = lijst(6, r)
End Function

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELACTIE {"
tekst = tekst & vbNewLine & vbTab & "ActieId = " & Me.ActieId
tekst = tekst & vbNewLine & vbTab & "MaterieelId = " & Me.MaterieelId
tekst = tekst & vbNewLine & vbTab & "Gebruiker = " & Me.Gebruiker
tekst = tekst & vbNewLine & vbTab & "Station = " & Me.Station
tekst = tekst & vbNewLine & vbTab & "Tijdstip = " & Me.Tijdstip
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Waarde = " & Me.Waarde
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByActieId()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    Set pr = cmd.CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.ActieId)), Me.ActieId)

    lijst = db.getLijstBySQL_AndParameter("SELECT * FROM MATERIEELACTIE WHERE ActieId = p1;", pr)

    FromList 0, lijst

End Function

Function InsertUitgifte() As Boolean
Me.Gebruiker = Environ$("username")
Me.Station = Environ$("computername")
Me.Tijdstip = Now()
Me.Omschrijving = "Uitgifte materieel vanuit magazijn"

If Me.insert = True Then InsertUitgifte = True
End Function

Function InsertInname() As Boolean
Me.Gebruiker = Environ$("username")
Me.Station = Environ$("computername")
Me.Tijdstip = Now()
Me.Omschrijving = "Inname materieel"

If Me.insert = True Then InsertInname = True
End Function

Function InsertBijwerkenVeld(Omschrijving As String) As Boolean
Dim tekst As Variant
tekst = Split(Omschrijving, "||")
Me.Gebruiker = Environ$("username")
Me.Station = Environ$("computername")
Me.Tijdstip = Now()
Me.Omschrijving = tekst(0)
Me.Waarde = tekst(1)
If Me.insert = True Then InsertBijwerkenVeld = True
End Function

