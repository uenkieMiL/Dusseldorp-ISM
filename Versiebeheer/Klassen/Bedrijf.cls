VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bedrijf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pBedrijfsnaam As String
Private pKVK As Long
Private pContactpersoon As String
Private pTelefoonnummer As String
Private pEmailadres As String

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
' Bedrijfsnaam property
'''''''''''''''''''''''''

Public Property Get Bedrijfsnaam() As String
    Bedrijfsnaam = pBedrijfsnaam
End Property

Public Property Let Bedrijfsnaam(Value As String)
    pBedrijfsnaam = Value
End Property

'''''''''''''''''''''''''
' KVK property
'''''''''''''''''''''''''

Public Property Get KVK() As Long
    KVK = pKVK
End Property

Public Property Let KVK(Value As Long)
    pKVK = Value
End Property

'''''''''''''''''''''''''
' Contactpersoon property
'''''''''''''''''''''''''

Public Property Get Contactpersoon() As String
    Contactpersoon = pContactpersoon
End Property

Public Property Let Contactpersoon(Value As String)
    pContactpersoon = Value
End Property

'''''''''''''''''''''''''
' Telefoonnummer property
'''''''''''''''''''''''''

Public Property Get Telefoonnummer() As String
    Telefoonnummer = pTelefoonnummer
End Property

Public Property Let Telefoonnummer(Value As String)
    pTelefoonnummer = Value
End Property

'''''''''''''''''''''''''
' Emailadres property
'''''''''''''''''''''''''

Public Property Get Emailadres() As String
    Emailadres = pEmailadres
End Property

Public Property Let Emailadres(Value As String)
    pEmailadres = Value
End Property

Public Function GetById(Id As Long)
    Dim lijst As Variant
    Dim strSQL As String
    Dim db As New DataBase
    
    strSQL = "SELECT * FROM BEDRIJVEN where Id = " & Id & ";"
    
    lijst = db.getLijstBySQL(strSQL)
    
    If IsEmpty(lijst) = False Then
        Me.Id = Id
        Me.KVK = lijst(1, 0)
        Me.Bedrijfsnaam = lijst(2, 0)
        If IsNull(lijst(3, 0)) = False Then Me.Contactpersoon = lijst(3, 0)
        If IsNull(lijst(4, 0)) = False Then Me.Telefoonnummer = lijst(4, 0)
        If IsNull(lijst(5, 0)) = False Then Me.Emailadres = lijst(5, 0)
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
    Dim db As New DataBase
    
    
 
    'update Statement
    strSQL = "UPDATE [BEDRIJVEN] " & _
             "SET " & _
             "[KVK] = ?, " & _
             "[Bedrijfsnaam] = ?, " & _
             "[Contactpersoon] = ?, " & _
             "[Telefoonnummer] = ?, " & _
             "[Email] = ? " & _
             "WHERE [Id] = ?;"
    
        db.Connect
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.KVK)), Me.KVK)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Bedrijfsnaam) + 1, Me.Bedrijfsnaam)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Contactpersoon) + 1, Me.Contactpersoon)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Telefoonnummer) + 1, Me.Telefoonnummer)
        .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Emailadres) + 1, Me.Emailadres)
        .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      update = False
    Else
      'Return a good return code
      update = True
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
    update = False
    Resume Exit_Insert
End Function

Public Function insert() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection

    'Instert Statement
    strSQL = "INSERT INTO [BEDRIJVEN] (" & _
                    "[KVK]," & _
                    "[Bedrijfsnaam]," & _
                    "[Contactpersoon]," & _
                    "[Telefoonnummer]," & _
                    "[Email]) " & vbCrLf & _
             "VALUES (p1,p2,p3,p4,p5);"
    
    'Define attachment to database table specifics
    
        
        con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adBigInt, adParamInput, Len(CStr(Me.KVK)), Me.KVK)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Bedrijfsnaam) + 1, Me.Bedrijfsnaam)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Contactpersoon) + 1, Me.Contactpersoon)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Telefoonnummer) + 1, Me.Telefoonnummer)
        .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Emailadres) + 1, Me.Emailadres)
        .CommandText = strSQL
        Set rst = .Execute
    End With
        rst.Open "SELECT @@identity AS NewID", con
        Me.Id = rst.Fields("NewID")
        
    'Return a good return code
    insert = True
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    con.Close
    Set con = Nothing
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    insert = False
    Resume Exit_Insert
End Function


Function ToString() As String
Dim tekst As String

tekst = "BEDRIJF {"
tekst = tekst & vbNewLine & "ID = " & Me.Id
tekst = tekst & vbNewLine & "Bedrijfsnaam = " & Me.Bedrijfsnaam
tekst = tekst & vbNewLine & "Contactpersoon = " & Me.Contactpersoon
tekst = tekst & vbNewLine & "Telefoonnummer = " & Me.Telefoonnummer
tekst = tekst & vbNewLine & "E-mailadres = " & Me.Emailadres
tekst = tekst & vbNewLine & "KVK-Nummer = " & Me.KVK
tekst = tekst & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function


