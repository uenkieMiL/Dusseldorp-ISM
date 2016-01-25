VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Uursoort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pOmschrijving As String
Private pKleur As Double
Private pKoppelbaar As Boolean
Private pCProjecten As Collection
Private pCPersoneel As Collection


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

Public Property Get Kleur() As Double
    Kleur = pKleur
End Property

Public Property Let Kleur(Value As Double)
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
' CProjecten property
'''''''''''''''''''''''''

Public Property Get CProjecten() As Collection
    Set CProjecten = pCProjecten
End Property

Public Property Let CProjecten(Value As Collection)
    Set pCProjecten = Value
End Property

'''''''''''''''''''''''''
' CProjecten property
'''''''''''''''''''''''''

Public Property Get CPersoneel() As Collection
    Set CPersoneel = pCPersoneel
End Property

Public Property Let CPersoneel(Value As Collection)
    Set pCPersoneel = Value
End Property


Property Get AantalProjecten() As Long
    AantalProjecten = pCProjecten.Count
End Property


Property Get AantalPersoneel() As Long
    AantalPersoneel = pCPersoneel.Count
End Property


Property Get PersoneelItem(ByVal vID As Variant) As Personeel
    Set PersoneelItem = pCPersoneel(vID)
End Property

Private Sub class_initialize()
    Set pCProjecten = New Collection
    Set pCPersoneel = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCProjecten = Nothing
    Set pCPersoneel = Nothing
End Sub

Public Sub ToevoegenProject(ByRef p As project)
    pCProjecten.Add p
End Sub

Public Sub ToevoegenPersoneel(ByRef p As Personeel)
    pCPersoneel.Add p
End Sub

Public Function GetById(Id As Long)
    Dim lijst As Variant
    Dim strSQL As String
   
    strSQL = "SELECT * FROM UURSOORT where Id = " & Id & ";"
    
    lijst = DataBase.LijstOpBasisVanQuery(strSQL)
    
    If IsEmpty(lijst) = False Then
        Me.Id = Id
        Me.Omschrijving = lijst(1, 0)
        Me.Kleur = lijst(2, 0)
        Me.Koppelbaar = lijst(3, 0)
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
    strSQL = "UPDATE [UURSOORT] " & _
             "SET " & _
             "[Omschrijving] = ?, " & _
             "[Kleur] = ?, " & _
             "[Koppelbaar] = ? " & _
             "WHERE [Id] = ?;"
    
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Omschrijving), Me.Omschrijving)
        .Parameters.Append .CreateParameter("p2", adDouble, adParamInput, Len(CStr(Me.Kleur)), Me.Kleur)
        .Parameters.Append .CreateParameter("p3", adBoolean, adParamInput, Len(CStr(Me.Koppelbaar)), Me.Koppelbaar)
        .Parameters.Append .CreateParameter("p4", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
    
        
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

Public Function insert() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection

    'Instert Statement
    strSQL = "INSERT INTO [UURSOORT] (" & _
                    "[Omschrijving], " & _
                    "[Kleur], " & _
                    "[Koppelbaar]) " & vbCrLf & _
             "VALUES (p1, p2, p3);"
    
    'Define attachment to database table specifics
    
        
        con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Omschrijving), Me.Omschrijving)
        .Parameters.Append .CreateParameter("p2", adDouble, adParamInput, Len(CStr(Me.Kleur)) + 1, Me.Kleur)
        .Parameters.Append .CreateParameter("p3", adBoolean, adParamInput, Len(CStr(Me.Koppelbaar)), Me.Koppelbaar)
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
Public Function delete() As Boolean
Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
     
    con.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
 
    'update Statement
    strSQL = "UPDATE [UURSOORT] " & _
             "SET " & _
             "[InActief] = ? " & _
             "WHERE [Id] = ?;"
    
        
        With cmd
        .ActiveConnection = con
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBoolean, adParamInput, 5, True)
        .Parameters.Append .CreateParameter("p3", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
        
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
Function ToString() As String
Dim tekst As String

tekst = "KLEUR {"
tekst = tekst & vbNewLine & vbTab & "ID = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Kleur = " & Me.Kleur
tekst = tekst & vbNewLine & vbTab & "Koppelbaar = " & Me.Koppelbaar
tekst = tekst & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function




