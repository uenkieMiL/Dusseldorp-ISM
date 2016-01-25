VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pSynergy As String
Private pVestiging As String
Private pSoort As Long
Private pStartDatum As Date
Private pEindDatum As Date
Private pStatus As Boolean
Private pCTaken As Collection
Private pTaakVanPlanning As taak


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
' Status property
'''''''''''''''''''''''''

Public Property Get Status() As Boolean
    Status = pStatus
End Property

Public Property Let Status(Value As Boolean)
    pStatus = Value
End Property
'''''''''''''''''''''''''
' CTaken property
'''''''''''''''''''''''''
Public Property Get cTaken() As Collection
  Set cTaken = pCTaken
End Property

Public Property Let cTaken(Value As Collection)
  Set pCTaken = Value
End Property

'''''''''''''''''''''''''
' TaakVanPlanning property
'''''''''''''''''''''''''

Public Property Get TaakVanPlanning() As taak
    
     Set TaakVanPlanning = pTaakVanPlanning
End Property
Public Property Let TaakVanPlanning(Value As taak)
    Set pTaakVanPlanning = Value
End Property

Property Get Count() As Long
    Count = pCTaken.Count
End Property

Property Get item(ByVal vID As Variant) As project
    Set item = pCTaken(vID)
End Property

Private Sub class_initialize()
    Set pCTaken = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCTaken = Nothing
End Sub

Public Sub ToevoegenTaak(ByRef t As taak)
    Dim at As taak
    Set at = New taak
    Set at = t
    pCTaken.Add t
End Sub

Function SoortByteNaarKortStringTerug() As String

    Select Case Me.soort
        Case 1
        SoortByteNaarKortStringTerug = "ACQ"
        
        Case 2
        SoortByteNaarKortStringTerug = "CALC"
        
        Case 3
        SoortByteNaarKortStringTerug = "WVB"
        
        Case 4
        SoortByteNaarKortStringTerug = "UITV"
    End Select
End Function

Function haalop(Id As Long)
Dim strSQL As String
Dim lijst As Variant
Dim db As New DataBase
lijst = db.getLijstBySQL("SELECT * FROM PLANNINGEN WHERE Id = " & Id)

If IsEmpty(lijst) = False Then
    Me.FromList 0, lijst
    Me.HaalTakenOp
End If
End Function

Function Create()
   Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim conn As New ADODB.connection
    Dim strSQL As String
    Dim db As New DataBase
    Dim maxl As Variant
    
    'Instert Statement
    strSQL = "INSERT INTO [PLANNINGEN] (" & _
                    "[Synergy], " & _
                    "[Vestiging], " & _
                    "[Soort], " & _
                    "[Startdatum], " & _
                    "[Einddatum]) " & _
             "VALUES (p1,p2,p3,p4,p5);"
    
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
        .CommandText = strSQL

        Set rst = .Execute
    End With
    
    'Return a good return code
    insert = True
    
    sql = "SELECT MAX([Id]) FROM [PLANNINGEN];"
    
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


Public Function update() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim con As New ADODB.connection
    Dim lRecordsAffected As Long
    Dim db As New DataBase
    
    
 
    'update Statement
    strSQL = "UPDATE [PLANNINGEN] " & _
             "SET " & _
             "[Synergy] = ?, " & _
             "[Vestiging] = ?, " & _
             "[Soort] = ?, " & _
             "[Startdatum] = ?, " & _
             "[Einddatum] = ?, " & _
             "[Status] = ? " & _
             "WHERE [Id] = ?;"
    
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
        .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Status)), Me.Status)
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


Function ToString() As String
Dim tekst As String

tekst = "Planning " & vbNewLine & "{"

tekst = tekst & vbNewLine & vbTab & "PlanningId = " & Me.Id
tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & vbNewLine & vbTab & "Soort = " & Me.soort
tekst = tekst & vbNewLine & vbTab & "Startdatum = " & Me.startdatum
tekst = tekst & vbNewLine & vbTab & "Einddatum = " & Me.einddatum
tekst = tekst & vbNewLine & vbTab & "Status  = " & Me.Status
tekst = tekst & vbNewLine & vbTab & "Aantal Taken in planning = " & Me.cTaken.Count

tekst = tekst & vbNewLine & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function

Function HaalTakenOp()
Dim lijst As Variant
Dim t As taak
Dim db As New DataBase

lijst = db.getLijstBySQL("SELECT id FROM TAKEN WHERE PlanningId = " & Me.Id & ";")

If IsEmpty(lijst) = False Then
    For x = 0 To UBound(lijst, 2)
        Set t = New taak
        t.haalop (lijst(0, x))
        Me.cTaken.Add t
    Next x
End If
End Function

Public Function FromList(r As Long, lijst As Variant)
    Me.Id = lijst(0, r)
    Me.synergy = lijst(1, r)
    Me.Vestiging = lijst(2, r)
    Me.soort = lijst(3, r)
    Me.startdatum = lijst(4, r)
    Me.einddatum = lijst(5, r)
    Me.Status = lijst(6, r)
End Function

Function VoegTakenToe()
Dim db As New DataBase
Dim t As taak
Dim lijst As Variant: lijst = db.getLijstBySQL("SELECT * FROM SJABLOON_TAKEN WHERE SOORT =" & Me.soort)
Dim datum As Date
Dim CKalender As New Collection
Set CKalender = Lijsten.KalenderOverallPlanning
    For r = 0 To UBound(lijst, 2)
        Set t = New taak
        t.planningid = Me.Id
        t.Omschrijving = lijst(4, r)
        t.Volgnummer = 1
        If lijst(6, r) = True Then datum = Me.startdatum Else datum = Me.einddatum
        If lijst(5, r) <> 0 Then datum = ZoekNieuweDatumVoorTaak(datum, CLng(lijst(5, r)), CKalender)
        t.startdatum = datum
        t.einddatum = datum
        t.Aantal = lijst(7, r)
        t.Ehd = lijst(8, r)
        t.veld = lijst(1, r)
        t.soort = lijst(2, r)
        t.BegrotingsRegel = False
        If PlanningSoort = 4 Then t.Bestekpost = True
        t.Opmerking = ""
        t.insert
    Next r
   
End Function

Function UpdateTakenNaarVerplaatsendatumPlanning(a As Integer)
Dim strSQL As String
Dim rst As New ADODB.Recordset
Dim db As New DataBase

db.Connect

strSQL = "SELECT * FROM TAKEN WHERE PlanningID = " & Me.Id & ";"

rst.Open Source:=strSQL, ActiveConnection:=db.connection, CursorType:=adOpenDynamic, LockType:=adLockOptimistic

If rst.BOF = False Or rst.EOF = False Then
    rst.MoveFirst
    Do Until rst.EOF = True
    
    rst.Fields("Startdatum") = DateAdd("d", a, rst.Fields("Startdatum"))
    rst.Fields("Einddatum") = DateAdd("d", a, rst.Fields("Einddatum"))
    rst.update
    rst.MoveNext
    Loop
    
    
End If
rst.Close
db.Disconnect

End Function

