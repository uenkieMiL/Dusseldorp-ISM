VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Productie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pSoort As Long
Private pOmschrijving As String
Private pStartDatum As Date
Private pEindDatum As Date
Private pKleur As Long
Private pGereed As Boolean
Private pSynergy As String
Private pVestiging As String
Private pFase As Byte


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
' Soort property
'''''''''''''''''''''''''

Public Property Get soort() As Long
    soort = pSoort
End Property

Public Property Let soort(Value As Long)
    pSoort = Value
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
' Kleur property
'''''''''''''''''''''''''

Public Property Get Kleur() As Long
    Kleur = pKleur
End Property

Public Property Let Kleur(Value As Long)
    pKleur = Value
End Property


'''''''''''''''''''''''''
' Gereed property
'''''''''''''''''''''''''

Public Property Get Gereed() As Boolean
    Gereed = pGereed
End Property

Public Property Let Gereed(Value As Boolean)
    pGereed = Value
End Property

'''''''''''''''''''''''''
' Fase property
'''''''''''''''''''''''''

Public Property Get Fase() As Byte
    Fase = pFase
End Property

Public Property Let Fase(Value As Byte)
    pFase = Value
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
' Vestiging property
'''''''''''''''''''''''''

Public Property Get Vestiging() As String
    Vestiging = pVestiging
End Property

Public Property Let Vestiging(Value As String)
    pVestiging = Value
End Property

Public Function update() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim db As New DataBase
    Dim con As ADODB.connection
    Dim lRecordsAffected As Long
     
    'update Statement
    strSQL = "UPDATE [PRODUCTIE] " & _
             "SET " & _
             "[Synergy] = ?," & _
             "[Vestiging] = ?," & _
             "[Soort] = ?," & _
             "[Startdatum] = ?," & _
             "[Einddatum] = ?," & _
             "[Kleur] = ?," & _
             "[Gereed] = ? " & _
             "WHERE [Id] = ?;"
    
    'Define attachment to database table specifics
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Vestiging), Me.Vestiging)
        .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.soort)), Me.soort)
        .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
        .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
        .Parameters.Append .CreateParameter("p6", adBigInt, adParamInput, Len(CStr(Me.Kleur)), Me.Kleur)
        .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.Gereed)), Me.Gereed)
        .Parameters.Append .CreateParameter("p8", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
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

Public Function insert() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim conn As New ADODB.connection
    Dim strSQL As String
    Dim db As New DataBase
    Dim maxl As Variant
    
    'Instert Statement
    strSQL = "INSERT INTO [PRODUCTIE] (" & _
                    "[Synergy], " & _
                    "[Vestiging], " & _
                    "[Soort], " & _
                    "[Startdatum], " & _
                    "[Einddatum], " & _
                    "[Kleur], " & _
                    "[Gereed]) " & _
             "VALUES (p1,p2,p3,p4,p5,p6,p7);"
    
    'Define attachment to database table specifics
    
        
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Vestiging), Me.Vestiging)
        .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, Len(CStr(Me.soort)), Me.soort)
        .Parameters.Append .CreateParameter("p4", adDate, adParamInput, Len(CStr(Me.startdatum)), Me.startdatum)
        .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.einddatum)), Me.einddatum)
        .Parameters.Append .CreateParameter("p6", adBigInt, adParamInput, Len(CStr(Me.Kleur)), Me.Kleur)
        .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.Gereed)), Me.Gereed)
        .CommandText = strSQL

        Set rst = .Execute
    End With
    
    'Return a good return code
    insert = True
    
    sql = "SELECT MAX([Id]) FROM [PRODUCTIE];"
            Set rs = New ADODB.Recordset
            rs.Open sql, db.connection, adOpenStatic, adLockReadOnly
            Me.Id = rs.Fields(0).Value
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    db.Disconnect
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    insert = False
    Resume Exit_Insert
End Function

Public Function haalop()
    
    Dim db As New DataBase
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
   
    strSQL = "SELECT * FROM PRODUCTIE where Id = " & Id & ";"
    
    db.Connect
    
    Set rst = db.OpenRecordset(strSQL)
    
    Me.Id = rst.Fields(0)
    Me.synergy = rst.Fields(1)
    Me.synergy = rst.Fields(2)
    Me.soort = rst.Fields(3)
    Me.startdatum = rst.Fields(4)
    Me.einddatum = rst.Fields(5)
    Me.Kleur = rst.Fields(6)
    Me.Gereed = rst.Fields(7)
    db.Disconnect
    
    Set rst = Nothing
    Set db = Nothing
    

End Function

Public Function delete() As Boolean
On Error GoTo ErrorHandler
Dim sql As String
Dim rst As ADODB.Recordset
Dim db As New DataBase

sql = "SELECT * FROM PRODUCTIE WHERE Id = " & Me.Id & ";"
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
