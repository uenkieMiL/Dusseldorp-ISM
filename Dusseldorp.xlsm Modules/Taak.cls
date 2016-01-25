VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Taak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pPlanningId As Long
Private pOmschrijving As String
Private pVolgnummer As Integer
Private pStartDatum As Date
Private pEindDatum As Date
Private pAantal As Double
Private pEhd As String
Private pStatus As Boolean
Private pOpmerking As String
Private pVeld As String
Private pSoort As String
Private pBegrotingsRegel As Boolean
Private pBestekpost As Boolean
Private pArtikelnummer As String
Private pCBehoeften As Collection


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
' PlanningId property
'''''''''''''''''''''''''
Public Property Get planningid() As Long
    planningid = pPlanningId
End Property

Public Property Let planningid(Value As Long)
    pPlanningId = Value
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
' Volgnummer property
'''''''''''''''''''''''''

Public Property Get Volgnummer() As Integer
    Volgnummer = pVolgnummer
End Property

Public Property Let Volgnummer(Value As Integer)
    pVolgnummer = Value
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
' Aantal property
'''''''''''''''''''''''''

Public Property Get Aantal() As Double
    Aantal = pAantal
End Property

Public Property Let Aantal(Value As Double)
    pAantal = Value
End Property

'''''''''''''''''''''''''
' Ehd property
'''''''''''''''''''''''''

Public Property Get Ehd() As String
    Ehd = pEhd
End Property

Public Property Let Ehd(Value As String)
    pEhd = Value
End Property

'''''''''''''''''''''''''
' Status property
'''''''''''''''''''''''''

Public Property Get Status() As Boolean
    Status = pStatus
End Property

Public Property Let Status(Value As Boolean)
    pStatus = Value
End Property

'''''''''''''''''''''''''
' Ehd property
'''''''''''''''''''''''''

Public Property Get Opmerking() As String
    Opmerking = pOpmerking
End Property

Public Property Let Opmerking(Value As String)
    pOpmerking = Value
End Property


'''''''''''''''''''''''''
' Veld property
'''''''''''''''''''''''''

Public Property Get veld() As String
    veld = pVeld
End Property

Public Property Let veld(Value As String)
    pVeld = Value
End Property

'''''''''''''''''''''''''
' Soort property
'''''''''''''''''''''''''

Public Property Get soort() As String
    soort = pSoort
End Property

Public Property Let soort(Value As String)
    pSoort = Value
End Property

'''''''''''''''''''''''''
' BegrotingsRegel property
'''''''''''''''''''''''''
Public Property Get BegrotingsRegel() As Boolean
    BegrotingsRegel = pBegrotingsRegel
End Property
Public Property Let BegrotingsRegel(Value As Boolean)
    pBegrotingsRegel = Value
End Property
'''''''''''''''''''''''''
' BegrotingsRegel property
'''''''''''''''''''''''''
Public Property Get Bestekpost() As Boolean
    Bestekpost = pBestekpost
End Property
Public Property Let Bestekpost(Value As Boolean)
    pBestekpost = Value
End Property
'''''''''''''''''''''''''
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Artikelnummer() As String
    Artikelnummer = pArtikelnummer
End Property

Public Property Let Artikelnummer(Value As String)
    pArtikelnummer = Value
End Property

'''''''''''''''''''''''''
' CBehoeften property
'''''''''''''''''''''''''

Public Property Get Cbehoeften() As Collection
  Set Cbehoeften = pCBehoeften
End Property

Public Property Let Cbehoeften(Value As Collection)
  Set pCBehoeften = Value
End Property
Private Sub class_initialize()
    Set pCBehoeften = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCBehoeften = Nothing
End Sub
Property Get Count() As Long
    Count = pCBehoeften.Count
End Property

Property Get item(ByVal vID As Variant) As project
    Set item = pCBehoeften(vID)
End Property
Public Sub ToevoegenBehoefte(ByRef bh As Behoefte)
    Dim b As New Behoefte
    Set b = bh
    pCBehoeften.Add b
End Sub

Function VoegTaakToe()
Dim db As New DataBase

Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
'Dim l As New Log
Dim logtekst As String

     cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT * FROM TAKEN WHERE 1=0"
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
        
    With rst
    .AddNew
    .Fields("PlanningId").Value = Me.planningid
    .Fields("Omschrijving").Value = Me.Omschrijving
    .Fields("Volgnummer").Value = Me.Volgnummer
    .Fields("Startdatum").Value = Me.startdatum
    .Fields("Einddatum").Value = Me.einddatum
    .Fields("Aantal").Value = Me.Aantal
    .Fields("Ehd").Value = Me.Ehd
    .Fields("Veld").Value = Me.veld
    .Fields("Soort").Value = Me.soort
    .Fields("Begrotingsregel").Value = Me.BegrotingsRegel
    .Fields("Bestekspost").Value = Me.Bestekpost
    If Me.Artikelnummer <> "" Then .Fields("Artikelnummer").Value = Me.Artikelnummer
    If Me.Opmerking <> "" Then .Fields("Opmerking").Value = Me.Opmerking
    .update
    Me.Id = .Fields("Id").Value
    .Close
    
    End With
    cnn.Close
    
    l.createLog Me.ToString, tk_aanmaken, Me.Id, taak

End Function

Function SoortByteNaarKortString(soort As Byte)
    Select Case soort
        Case 1
        Me.soort = "ACQ"
        
        Case 2
        Me.soort = "CALC"
        
        Case 3
        Me.soort = "WVB"
        
        Case 4
        Me.soort = "UITV"
    End Select
End Function


Function haalop(Id As Long)
Dim strSQL As String
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
Dim db As New DataBase

db.Connect

strSQL = "SELECT * FROM TAKEN WHERE Id = " & Id

rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenDynamic, LockType:=adLockOptimistic

Me.Id = Id
Me.Omschrijving = rst.Fields("Omschrijving")
Me.Volgnummer = rst.Fields("Volgnummer")
Me.startdatum = rst.Fields("Startdatum")
Me.einddatum = rst.Fields("Einddatum")
Me.Aantal = rst.Fields("Aantal")
Me.Ehd = rst.Fields("Ehd")
Me.Status = rst.Fields("Status")
Me.veld = rst.Fields("veld")
Me.soort = rst.Fields("Soort")
If IsNull(rst.Fields("Opmerking")) = True Then Me.Opmerking = "" Else Me.Opmerking = rst.Fields("Opmerking")
Me.planningid = rst.Fields("PlanningId")
Me.BegrotingsRegel = rst.Fields("Begrotingsregel")
Me.Bestekpost = rst.Fields("Bestekspost")

rst.Close
db.Disconnect
End Function

Public Function Verwijder()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
'Dim l As New Log

Dim datum As Date

 cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT TOP 1 * FROM TAKEN WHERE ID=" & Me.Id
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
    
    rst.delete
    
    rst.Close
    cnn.Close
    l.createLog Me.ToString & vbNewLine & "VERWIJDERD", tk_vewijderen, Me.Id, taak
    
End Function

Public Function delete() As Boolean
On Error GoTo ErrorHandler
Dim sql As String
Dim rst As ADODB.Recordset
Dim db As New DataBase

sql = "SELECT * FROM TAKEN WHERE Id = " & Me.Id & ";"
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

Public Function update() As Boolean
On Error GoTo Err_Update
  Dim adoCMD As New ADODB.Command
  Dim adoRS As New ADODB.Recordset
  Dim strSQL As String
  Dim lRecordsAffected As Long
  Dim db As New DataBase
  

  'Define a query to Update a new record into the FE temp table
  strSQL = "UPDATE TAKEN " & vbCrLf & _
           "SET Planningid = ?," & vbCrLf & _
           "Omschrijving = ?," & vbCrLf & _
           "Volgnummer= ?," & vbCrLf & _
           "Startdatum = ?," & vbCrLf & _
           "Einddatum = ?," & vbCrLf & _
           "Aantal = ?," & vbCrLf & _
           "Ehd = ?," & vbCrLf & _
           "Status = ?," & vbCrLf & _
           "Veld = ?," & vbCrLf & _
           "Soort = ?," & vbCrLf & _
           "Begrotingsregel = ?," & vbCrLf & _
           "Opmerking = ?," & vbCrLf & _
           "Artikelnummer = ?," & vbCrLf & _
           "Bestekspost= ?" & vbCrLf & _
           "WHERE TAKEN.Id = ?;"

    db.Connect


  'Define attachment to database table specifics
  With adoCMD
    .ActiveConnection = db.connection
    .CommandType = adCmdText
    .CommandText = strSQL
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, 5, Me.planningid)
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, 255, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p4", adInteger, adParamInput, 4, Me.Volgnummer)
    .Parameters.Append .CreateParameter("p5", adDate, adParamInput, 8, Me.startdatum)
    .Parameters.Append .CreateParameter("p6", adDate, adParamInput, 8, Me.einddatum)
    .Parameters.Append .CreateParameter("p7", adDouble, adParamInput, 6, Me.Aantal)
    .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, 3, Me.Ehd)
    .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, 1, Me.Status)
    .Parameters.Append .CreateParameter("p10", adVarChar, adParamInput, 10, Me.veld)
    .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, 10, Me.soort)
    .Parameters.Append .CreateParameter("p12", adBoolean, adParamInput, 1, Me.BegrotingsRegel)
    .Parameters.Append .CreateParameter("p13", adVarChar, adParamInput, 255, Me.Opmerking)
    .Parameters.Append .CreateParameter("p14", adVarChar, adParamInput, 255, Me.Artikelnummer)
    .Parameters.Append .CreateParameter("p15", adBoolean, adParamInput, 1, Me.Bestekpost)
    .Parameters.Append .CreateParameter("p16", adInteger, adParamInput, 6, Me.Id)
    
    Set adoRS = .Execute(lRecordsAffected)
  End With

  If lRecordsAffected = 0 Then
    update = False
  Else
    'Return a good return code
    update = True
  End If

Exit_Update:
  'Clean up the connection to the database
  Set adoCMD = Nothing
  Set adoRS = Nothing

  Exit Function

Err_Update:
 'Call errorhandler_MsgBox("Class: clsObjPartsImportWizardTbl, Function: Update()")
 'Update = False
Resume Exit_Update

End Function

Function ToString() As String
Dim tekst As String

tekst = "TAAK " & vbNewLine & "{"
tekst = tekst & vbNewLine & vbTab & "ID = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "PlanningId = " & Me.planningid
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Volgnummer  = " & Me.Volgnummer
tekst = tekst & vbNewLine & vbTab & "Startdatum = " & Me.startdatum
tekst = tekst & vbNewLine & vbTab & "Einddatum = " & Me.einddatum
tekst = tekst & vbNewLine & vbTab & "Aantal = " & Me.Aantal
tekst = tekst & vbNewLine & vbTab & "Ehd = " & Me.Ehd
tekst = tekst & vbNewLine & vbTab & "Status = " & Me.Status
tekst = tekst & vbNewLine & vbTab & "Veld = " & Me.veld
tekst = tekst & vbNewLine & vbTab & "Soort = " & Me.soort
tekst = tekst & vbNewLine & vbTab & "BegrotingsRegel = " & Me.BegrotingsRegel
tekst = tekst & vbNewLine & vbTab & "Bestekspost = " & Me.Bestekpost
tekst = tekst & vbNewLine & vbTab & "Artikelnummer = " & Me.Artikelnummer

tekst = tekst & vbNewLine & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function

Function FromList(r As Long, lijst As Variant)
    Me.Id = lijst(0, r)
    Me.planningid = lijst(1, r)
    Me.Omschrijving = lijst(2, r)
    Me.Volgnummer = lijst(3, r)
    Me.startdatum = lijst(4, r)
    Me.einddatum = lijst(5, r)
    Me.Aantal = lijst(6, r)
    Me.Ehd = lijst(7, r)
    Me.Status = lijst(8, r)
    Me.veld = lijst(9, r)
    Me.soort = lijst(10, r)
    Me.BegrotingsRegel = lijst(11, r)
    If IsNull(lijst(12, r)) = False Then Me.Opmerking = lijst(12, r)
    If IsNull(lijst(13, r)) = False Then Me.Artikelnummer = lijst(13, r)
    Me.Bestekpost = lijst(14, r)
End Function

Function CopyTaak()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset
'Dim l As New Log
Dim db As New DataBase
Dim datum As Date

db.Connect
    
    strSQL = "SELECT * FROM TAKEN WHERE 1=0"
    
    rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
        
    With rst
    .AddNew
    .Fields("PlanningId").Value = Me.planningid
    .Fields("Omschrijving").Value = Me.Omschrijving
    .Fields("Volgnummer").Value = Me.Volgnummer + 1
    .Fields("Startdatum").Value = Me.startdatum
    .Fields("Einddatum").Value = Me.einddatum
    .Fields("Aantal").Value = Me.Aantal
    .Fields("Ehd").Value = Me.Ehd
    .Fields("Veld").Value = Me.veld
    .Fields("Soort").Value = Me.soort
    .Fields("Begrotingsregel").Value = Me.BegrotingsRegel
    If Me.Opmerking <> "" Then .Fields("Opmerking").Value = Me.Opmerking
    .update
    Me.Id = rst.Fields("Id")
    Me.Volgnummer = rst.Fields("Volgnummer")
    .Close
    
    
    End With
    db.Disconnect

'l.createLog t.ToString, tk_aanmaken, t.id, Taak
End Function


Public Function insert() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim conn As New ADODB.connection
    Dim strSQL As String
    Dim db As New DataBase
    Dim maxl As Variant
    
    'Instert Statement
    strSQL = "INSERT INTO [TAKEN] (" & _
                    "[PlanningId], " & _
                    "[Omschrijving], " & _
                    "[Volgnummer], " & _
                    "[Startdatum], " & _
                    "[Einddatum], " & _
                    "[Aantal], " & _
                    "[Ehd], " & _
                    "[Status], " & _
                    "[Veld], " & _
                    "[Soort], " & _
                    "[Begrotingsregel], " & _
                    "[Opmerking], " & _
                    "[Artikelnummer], " & _
                    "[Bestekspost]) " & _
             "VALUES (p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,p13,p14);"
    
    'Define attachment to database table specifics
    
        
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, 5, Me.planningid)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, 255, Me.Omschrijving)
        .Parameters.Append .CreateParameter("p3", adInteger, adParamInput, 4, Me.Volgnummer)
        .Parameters.Append .CreateParameter("p4", adDate, adParamInput, 8, Me.startdatum)
        .Parameters.Append .CreateParameter("p5", adDate, adParamInput, 8, Me.einddatum)
        .Parameters.Append .CreateParameter("p6", adDouble, adParamInput, 6, Me.Aantal)
        .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, 3, Me.Ehd)
        .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, 1, Me.Status)
        .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, 10, Me.veld)
        .Parameters.Append .CreateParameter("p10", adVarChar, adParamInput, 10, Me.soort)
        .Parameters.Append .CreateParameter("p11", adBoolean, adParamInput, 1, Me.BegrotingsRegel)
        .Parameters.Append .CreateParameter("p12", adVarChar, adParamInput, 255, Me.Opmerking)
        .Parameters.Append .CreateParameter("p13", adVarChar, adParamInput, 255, Me.Artikelnummer)
        .Parameters.Append .CreateParameter("p14", adBoolean, adParamInput, 1, Me.Bestekpost)
        .CommandText = strSQL

        Set rst = .Execute
    End With
    
    'Return a good return code
    insert = True
    
    sql = "SELECT MAX([Id]) FROM [TAKEN];"
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

