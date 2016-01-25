VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Datum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const tabelnaam As String = "KALENDER"

Private pDatum As Date
Private pKolomnummer As Long
Private pFeestdag As Boolean
Private pExtraDag As Boolean
Private pZichtbaar As Boolean
Private pOmschrijving As String

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
' ExtraDag property
'''''''''''''''''''''''''

Public Property Get ExtraDag() As Boolean
ExtraDag = pExtraDag
End Property

Public Property Let ExtraDag(Value As Boolean)
    pExtraDag = Value
End Property

'''''''''''''''''''''''''
' Feestdag property
'''''''''''''''''''''''''

Public Property Get feestdag() As Boolean
feestdag = pFeestdag
End Property

Public Property Let feestdag(Value As Boolean)
    pFeestdag = Value
End Property

'''''''''''''''''''''''''
' Zichtbaar property
'''''''''''''''''''''''''

Public Property Get Zichtbaar() As Boolean
Zichtbaar = pZichtbaar
End Property

Public Property Let Zichtbaar(Value As Boolean)
    pZichtbaar = Value
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


Public Function update() As Boolean
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
    Dim db As New DataBase
    Dim con As ADODB.connection
    Dim lRecordsAffected As Long
     
    'update Statement
    strSQL = "UPDATE [" & tabelnaam & "] " & _
             "SET " & _
             "[Feestdag] = ?," & _
             "[Extradag] = ?," & _
             "[Omschrijving] = ?," & _
             "[Zichtbaar] = ? " & _
             "WHERE [Datum] = ?;"
    
    'Define attachment to database table specifics
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBoolean, adParamInput, Len(CStr(Me.feestdag)), Me.feestdag)
        .Parameters.Append .CreateParameter("p2", adBoolean, adParamInput, Len(CStr(Me.ExtraDag)), Me.ExtraDag)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
        .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Zichtbaar)), Me.Zichtbaar)
        .Parameters.Append .CreateParameter("p5", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
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
    Dim conn As New ADODB.connection
    Dim strSQL As String
    Dim db As New DataBase
    
    'Instert Statement
    strSQL = "INSERT INTO [" & tabelnaam & "] (" & _
             "[Datum], " & _
             "[Feestdag], " & _
             "[Extradag], " & _
             "[Omschrijving], " & _
             "[Zichtbaar]) " & _
             "VALUES (p1,p2,p3,p4,p5);"
    
    'Define attachment to database table specifics
    
        
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adDate, adParamInput, Len(CStr(Me.datum)), Me.datum)
        .Parameters.Append .CreateParameter("p2", adBoolean, adParamInput, Len(CStr(Me.feestdag)), Me.feestdag)
        .Parameters.Append .CreateParameter("p3", adBoolean, adParamInput, Len(CStr(Me.ExtraDag)), Me.ExtraDag)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
        .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Zichtbaar)), Me.Zichtbaar)
        
        .CommandText = strSQL

        Set rst = .Execute
    End With
    
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

Public Function readbydate()
    
    Dim db As New DataBase
    Dim rst As New ADODB.Recordset
    Dim strSQL As String
   
    strSQL = "SELECT * FROM " & tabelnaam & " where Datum = #" & Year(Me.datum) & "-" & Month(Me.datum) & "-" & Day(Me.datum) & "#;"
    
    db.Connect
    
    Set rst = db.OpenRecordset(strSQL)
    
    If rst.BOF = False Or rst.EOF = False Then
        Me.datum = rst.Fields(0)
        Me.feestdag = rst.Fields(1)
        Me.ExtraDag = rst.Fields(2)
        Me.Omschrijving = rst.Fields(3)
        Me.Zichtbaar = rst.Fields(4)
    Else
        Functies.errorhandler_MsgBox ("Er is geen datum gevonden")
    End If
    
    db.Disconnect
    Set rst = Nothing
    Set db = Nothing
    

End Function

Public Function delete() As Boolean
On Error GoTo ErrorHandler
Dim sql As String
Dim rst As ADODB.Recordset
Dim db As New DataBase

sql = "SELECT * FROM " & tabelnaam & " WHERE Datum = #" & Year(Me.datum) & "-" & Month(Me.datum) & "-" & Day(Me.datum) & "#;"
'We are using a select statement that will return only

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

Function ToString() As String
Dim tekst As String

tekst = tabelnaam & " {"
tekst = tekst & vbNewLine & "Datum = " & Me.datum
tekst = tekst & vbNewLine & "Feestdag = " & Me.feestdag
tekst = tekst & vbNewLine & "Extradag = " & Me.ExtraDag
tekst = tekst & vbNewLine & "Omschrijving = " & Me.Omschrijving
tekst = tekst & vbNewLine & "Zichtbaar = " & Me.Zichtbaar
tekst = tekst & vbNewLine & "Kolomnummer = " & Me.Kolomnummer
tekst = tekst & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function

