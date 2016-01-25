VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Behoefte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pPlaningId As Long
Private pTaakId As Long
Private pDatum As Date
Private pHoeveelheid As Double
Private pSynergy As String
Private pKolomnummer As Long

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
' PlaningId property
'''''''''''''''''''''''''

Public Property Get planingid() As Long
    planingid = pPlaningId
End Property

Public Property Let planingid(Value As Long)
    pPlaningId = Value
End Property

'''''''''''''''''''''''''
' TaakId property
'''''''''''''''''''''''''

Public Property Get taakid() As Long
    taakid = pTaakId
End Property

Public Property Let taakid(Value As Long)
    pTaakId = Value
End Property

'''''''''''''''''''''''''
' Datum property
'''''''''''''''''''''''''

Public Property Get datum() As Date
    datum = pDatum
End Property

Public Property Let datum(Value As Date)
    pDatum = Value
    pKolomnummer = DatumNaarKolom(Value)
End Property

'''''''''''''''''''''''''
' Hoeveelheid property
'''''''''''''''''''''''''

Public Property Get hoeveelheid() As Double
    hoeveelheid = pHoeveelheid
End Property

Public Property Let hoeveelheid(Value As Double)
    pHoeveelheid = Value
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
' Kolomnummer property
'''''''''''''''''''''''''

Public Property Get Kolomnummer() As Long
    Kolomnummer = pKolomnummer
End Property

Function DatumNaarKolom(datum As Date) As Long
Dim a As Long
Dim gelukt As Boolean
Dim rng As Range
Set rng = ThisWorkbook.Sheets(Blad3.Name).Range("A10")
Waarde = rng.Offset(a)
Do While Waarde <> ""

If Waarde = datum Then
    DatumNaarKolom = a
    gelukt = True
    Exit Do
End If

    a = a + 1
Waarde = rng.Offset(a)

Loop

If gelukt = False Then DatumNaarKolom = -1
End Function

Public Function VoegToe()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

Dim datum As Date

 cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT * FROM BEHOEFTEN WHERE 1=0"
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
        
    With rst
        .AddNew
        .Fields("TaakId").Value = Me.taakid
        .Fields("PlanningId").Value = Me.planingid
        .Fields("Synergy").Value = Me.synergy
        .Fields("datum").Value = Me.datum
        .Fields("Hoeveelheid").Value = Me.hoeveelheid
        .update
    Me.Id = rst.Fields("Id")
        .Close
       
    End With
    cnn.Close

End Function


Public Function Verwijder()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

Dim datum As Date

 cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT TOP 1 * FROM BEHOEFTEN WHERE ID=" & Me.Id
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
    
    rst.delete
    
    rst.Close
    cnn.Close

End Function

Public Function UpdateAantal()
Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

Dim datum As Date

 cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT TOP 1 * FROM BEHOEFTEN WHERE ID=" & Me.Id
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
    
    rst.Fields("Hoeveelheid") = Me.hoeveelheid
    rst.update
    
    rst.Close
    cnn.Close

End Function

Public Function haalop(datum As Date, taakid As Long, synergy As String, planningid As Long) As Boolean

Dim cnn As New ADODB.connection
Dim rst As New ADODB.Recordset

 cnn.Open ConnectionString:="Provider = Microsoft.ACE.OLEDB.12.0; data source=" & myconn
    
    strSQL = "SELECT TOP 1 * FROM BEHOEFTEN WHERE Synergy = '" & synergy & "' AND TaakId=" & taakid & _
    " AND PlanningId =" & planningid & " AND Datum = #" & Format(datum, "yyyy-mm-dd") & "#"
    
    rst.Open Source:=strSQL, ActiveConnection:=cnn, CursorType:=adOpenKeyset, LockType:=adLockPessimistic, Options:=adCmdText
    If rst.EOF = False And rst.BOF = False Then
    Me.Id = rst.Fields("Id")
    Me.synergy = rst.Fields("Synergy")
    Me.planingid = rst.Fields("PlanningId")
    Me.taakid = rst.Fields("TaakId")
    Me.datum = rst.Fields("datum")
    Me.hoeveelheid = rst.Fields("Hoeveelheid")
    haalop = True
    End If
    
    rst.Close
    cnn.Close
End Function

