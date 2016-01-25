VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersoneelPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pPersoneelId As Long
Private pUursoortId As Long
Private pUursoort As New Uursoort
Private pDatum As Date
Private pKolomnummer As Long
Private pSynergy As String

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
' PersoneelId property
'''''''''''''''''''''''''

Public Property Get personeelid() As Long
    personeelid = pPersoneelId
End Property

Public Property Let personeelid(Value As Long)
    pPersoneelId = Value
End Property

'''''''''''''''''''''''''
' UursoortId property
'''''''''''''''''''''''''

Public Property Get UursoortId() As Long
    UursoortId = pUursoortId
End Property

Public Property Let UursoortId(Value As Long)
    pUursoortId = Value
End Property

'''''''''''''''''''''''''
' Datum property
'''''''''''''''''''''''''

Public Property Get datum() As Date
    datum = pDatum
End Property

Public Property Let datum(Value As Date)
    pDatum = Value
End Property

'''''''''''''''''''''''''
' Kolomnummer property
'''''''''''''''''''''''''

Public Property Get Kolomnummer() As Long
    Kolomnummer = pKolomnummer
End Property

Public Property Let Kolomnummer(Value As Long)
    pKolomnummer = Value
End Property

'''''''''''''''''''''''''
' Uursoort property
'''''''''''''''''''''''''

Public Property Get Uursoort() As Uursoort
   Set Uursoort = pUursoort
End Property

Public Property Let Uursoort(Value As Uursoort)
   Set pUursoort = Value
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
Public Function GetById(Id As Long)
    Dim lijst As Variant
    Dim strSQL As String
   
    strSQL = "SELECT * FROM PLANNING_PERSONEEL where Id = " & Id & ";"
    
    lijst = DataBase.LijstOpBasisVanQuery(strSQL)
    
    If IsEmpty(lijst) = False Then
        Me.Id = Id
        Me.personeelid = lijst(1, 0)
        Me.datum = lijst(2, 0)
        Me.UursoortId = lijst(3, 0)
        Me.UursoortId = lijst(4, 0)
    End If
    
End Function

Public Sub save()
    If pId <> 0 Then
        update
    Else
        'id is onbekend, calculatie bestaat niet -> aanmaken
        insert
    End If
End Sub
Private Function update() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
     
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
 
    'update Statement
    strSQL = "UPDATE [PLANNING_PERSONEEL] " & _
             "SET " & _
             "[PersoneelId] = ?, " & _
             "[Datum] = ?, " & _
             "[UursoortId] = ?, " & _
             "[Synergy] = ? " & _
             "WHERE [Id] = ?;"
    
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        .Parameters.Append .CreateParameter("p3", adBigInt, adParamInput, Len(CStr(Me.UursoortId)), Me.UursoortId)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
        .Parameters.Append .CreateParameter("p5", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
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
    Set con = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    update = False
    Resume Exit_Insert
End Function

Public Function updateDatumPersoneel() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
     
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
 
    'update Statement
    strSQL = "UPDATE [PLANNING_PERSONEEL] " & _
             "SET " & _
             "[UursoortId] = ?, " & _
             "[Synergy] = ? " & _
             "WHERE [DATUM] = ? AND [PersoneelId] = ?;"
    
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.UursoortId)), Me.UursoortId)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
        .Parameters.Append .CreateParameter("p3", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        .Parameters.Append .CreateParameter("p4", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      updateDatumPersoneel = False
    Else
      'Return a good return code
      updateDatumPersoneel = True
    End If
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set con = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    updateDatumPersoneel = False
    Resume Exit_Insert
End Function

Public Function insert() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim db As New DataBase
    
    'Instert Statement
    strSQL = "INSERT INTO [PLANNING_PERSONEEL] (" & _
                    "[PersoneelId], " & _
                    "[Datum], " & _
                    "[UursoortId], " & _
                    "[Synergy]) " & vbCrLf & _
             "VALUES (p1, p2, p3, p4);"
    
    'Define attachment to database table specifics
    
        
       db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.datum)) + 1, Me.datum)
        .Parameters.Append .CreateParameter("p3", adBigInt, adParamInput, Len(CStr(Me.UursoortId)), Me.UursoortId)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.synergy) + 1, Me.synergy)
        .CommandText = strSQL
        Set rst = .Execute
    End With
        rst.Open "SELECT @@identity AS NewID", db.connection
        Me.Id = rst.Fields("NewID")
        
    'Return a good return code
    insert = True
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    db.Disconnect
    Set db = Nothing
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    insert = False
    Resume Exit_Insert
End Function
Public Function delete() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
     
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_PERSONEEL] " & _
             "WHERE [Id] = ?;"
    
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      delete = False
    Else
      'Return a good return code
      delete = True
    End If
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set con = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    delete = False
    Resume Exit_Insert
End Function
Public Function DeleteDatumPersoneel() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
    db.Connect
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_PERSONEEL] " & _
             "WHERE [PersoneelId] = ? AND [Datum] = ?;"
    
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    db.Disconnect
    
    If lRecordsAffected = 0 Then
      DeleteDatumPersoneel = False
    Else
      'Return a good return code
      DeleteDatumPersoneel = True
    End If
    
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumPersoneel = False
    Resume Exit_Insert

End Function

Public Function DeleteDatumPersoneelUursoort() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
     
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_PERSONEEL] " & _
             "WHERE [PersoneelId] = ? AND [Datum] = ? AND [UursoortId] = ?;"
    
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        .Parameters.Append .CreateParameter("p3", adBigInt, adParamInput, Len(CStr(Me.UursoortId)), Me.UursoortId)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      DeleteDatumPersoneelUursoort = False
    Else
      'Return a good return code
      DeleteDatumPersoneelUursoort = True
    End If
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set con = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumPersoneelUursoort = False
    Resume Exit_Insert
End Function

Public Function DeleteDatumPersoneelSynergy() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
    db.Connect
 
    'update Statement
    strSQL = "DELETE FROM [PLANNING_PERSONEEL] " & _
             "WHERE [PersoneelId] = ? AND [Datum] = ? AND [Synergy] = ? AND [UursoortId] = ?;"
    
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.personeelid)), Me.personeelid)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p4", adBigInt, adParamInput, Len(CStr(Me.UursoortId)), Me.UursoortId)
        Set rst = .Execute(lRecordsAffected)
    End With
    
    db.Disconnect
    
    If lRecordsAffected = 0 Then
      DeleteDatumPersoneelSynergy = False
    Else
      'Return a good return code
      DeleteDatumPersoneelSynergy = True
    End If
    
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    Set db = Nothing
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    DeleteDatumPersoneelSynergy = False
    Resume Exit_Insert
End Function


Function ToString() As String
Dim tekst As String

tekst = "PersoneelsPlanning {"
tekst = tekst & vbNewLine & vbTab & "ID = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "PersoneelId = " & Me.personeelid
tekst = tekst & vbNewLine & vbTab & "Datum = " & Me.datum
tekst = tekst & vbNewLine & vbTab & "UursoortId = " & Me.UursoortId
tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function


