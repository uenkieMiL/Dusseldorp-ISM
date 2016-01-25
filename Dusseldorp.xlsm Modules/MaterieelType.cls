VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterieelType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pMaterieeltypeid As Long
Private pArtikelnummer As String
Private pOmschrijving As String
Private pInactief As Boolean




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
' Artikelnummer property
'''''''''''''''''''''''''

Public Property Get Artikelnummer() As String
    Artikelnummer = pArtikelnummer
End Property

Public Property Let Artikelnummer(Value As String)
    pArtikelnummer = Value
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
' Inactief property
'''''''''''''''''''''''''

Public Property Get Inactief() As Boolean
    Inactief = pInactief
End Property

Public Property Let Inactief(Value As Boolean)
    pInactief = Value
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
strSQL = "INSERT INTO [MATERIEELTYPEN] (" & _
            "[MaterieelTypeId], " & _
            "[Artikelnummer], " & _
            "[Omschrijving], " & _
            "[Inactief]) " & _
"VALUES (p1, p2, p3, p4);"

'Define attachment to database table specifics

With cmd
    .ActiveConnection = db.connection
    .CommandType = adCmdText

    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelTypeId)), Me.MaterieelTypeId)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Artikelnummer) + 1, Me.Artikelnummer)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
    .CommandText = strSQL

    Set rst = .Execute
End With

'Return a good return code
insert = True

strSQL = "SELECT MAX([MaterieelTypeId]) FROM [MATERIEELTYPEN];"
Set rs = New ADODB.Recordset
rs.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
Me.MaterieelTypeId = rs.Fields(0).Value
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
strSQL = "UPDATE [MATERIEELTYPEN] " & _
         "SET " & _
         "[Artikelnummer] = ?, " & _
         "[Omschrijving] = ?, " & _
         "[Inactief] = ? " & _
         "WHERE [MaterieelTypeId] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL

    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Artikelnummer) + 1, Me.Artikelnummer)
    .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.MaterieelTypeId)), Me.MaterieelTypeId)

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

sql = "SELECT * FROM MATERIEELTYPEN WHERE MaterieelTypeId = " & Me.MaterieelTypeId & ";"
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
Function FromList(r As Long, lijst As Variant, Optional V As Long = 0)
    If IsEmpty(lijst) = False Then
        Me.MaterieelTypeId = lijst(0 + V, r)
        Me.Artikelnummer = lijst(1 + V, r)
        Me.Omschrijving = lijst(2 + V, r)
        Me.Inactief = lijst(3 + V, r)
    End If
End Function

Function ToString() As String
Dim tekst As String
tekst = "MATERIEELTYPE {"
tekst = tekst & vbNewLine & vbTab & "MaterieelTypeId = " & Me.MaterieelTypeId
tekst = tekst & vbNewLine & vbTab & "Artikelnummer = " & Me.Artikelnummer
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Inactief = " & Me.Inactief
tekst = tekst & vbNewLine & "}"
ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function
Function GetByMaterieelTypeId()
Dim db As New DataBase
Dim pr As New ADODB.Parameter
Dim cmd As New ADODB.Command
Dim lijst As Variant

    lijst = db.getLijstBySQL("SELECT * FROM MATERIEELTYPEN WHERE MaterieelTypeId = " & Me.MaterieelTypeId & ";")
    
    If IsEmpty(lijst) = False Then
        FromList 0, lijst
    End If
    
    
    
End Function

