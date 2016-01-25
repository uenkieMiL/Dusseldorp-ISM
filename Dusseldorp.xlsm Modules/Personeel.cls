VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Personeel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pId As Long
Private pAchternaam As String
Private Pnaam As String
Private pBSN As Long
Private pMachinist As Boolean
Private pTimmerman As Boolean
Private pGrondwerker As Boolean
Private pSloper As Boolean
Private pDHV As Boolean
Private pDTA As Boolean
Private pUitvoerder As Boolean
Private pKVP As Boolean
Private pHVK As Boolean
Private pBijzonderheden As String
Private pBeoordeling As Integer
Private pArchief As Boolean
Private pBedrijfId As Long
Private pBedrijf As New Bedrijf
Private pCPersoneelPlanning As Collection



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
' Achternaam property
'''''''''''''''''''''''''

Public Property Get Achternaam() As String
    Achternaam = pAchternaam
End Property

Public Property Let Achternaam(Value As String)
    pAchternaam = Value
End Property

'''''''''''''''''''''''''
' Naam property
'''''''''''''''''''''''''

Public Property Get Naam() As String
    Naam = Pnaam
End Property

Public Property Let Naam(Value As String)
    Pnaam = Value
End Property

'''''''''''''''''''''''''
' BSN property
'''''''''''''''''''''''''

Public Property Get BSN() As Long
    BSN = pBSN
End Property

Public Property Let BSN(Value As Long)
    pBSN = Value
End Property

'''''''''''''''''''''''''
' Machinist property
'''''''''''''''''''''''''

Public Property Get Machinist() As Boolean
    Machinist = pMachinist
End Property

Public Property Let Machinist(Value As Boolean)
    pMachinist = Value
End Property

'''''''''''''''''''''''''
' Timmerman property
'''''''''''''''''''''''''

Public Property Get Timmerman() As Boolean
    Timmerman = pTimmerman
End Property

Public Property Let Timmerman(Value As Boolean)
    pTimmerman = Value
End Property

'''''''''''''''''''''''''
' Grondwerker property
'''''''''''''''''''''''''

Public Property Get Grondwerker() As Boolean
    Grondwerker = pGrondwerker
End Property

Public Property Let Grondwerker(Value As Boolean)
    pGrondwerker = Value
End Property

'''''''''''''''''''''''''
' Sloper property
'''''''''''''''''''''''''

Public Property Get Sloper() As Boolean
    Sloper = pSloper
End Property

Public Property Let Sloper(Value As Boolean)
    pSloper = Value
End Property

'''''''''''''''''''''''''
' DHV property
'''''''''''''''''''''''''

Public Property Get DHV() As Boolean
    DHV = pDHV
End Property

Public Property Let DHV(Value As Boolean)
    pDHV = Value
End Property

'''''''''''''''''''''''''
' KVP property
'''''''''''''''''''''''''

Public Property Get KVP() As Boolean
    KVP = pKVP
End Property

Public Property Let KVP(Value As Boolean)
    pKVP = Value
End Property

'''''''''''''''''''''''''
' DHV property
'''''''''''''''''''''''''

Public Property Get HVK() As Boolean
    HVK = pHVK
End Property

Public Property Let HVK(Value As Boolean)
    pHVK = Value
End Property

'''''''''''''''''''''''''
' DTA property
'''''''''''''''''''''''''

Public Property Get DTA() As Boolean
    DTA = pDTA
End Property

Public Property Let DTA(Value As Boolean)
    pDTA = Value
End Property

'''''''''''''''''''''''''
' Uitvoerder property
'''''''''''''''''''''''''

Public Property Get Uitvoerder() As Boolean
    Uitvoerder = pUitvoerder
End Property

Public Property Let Uitvoerder(Value As Boolean)
    pUitvoerder = Value
End Property

'''''''''''''''''''''''''
' Bijzonderheden property
'''''''''''''''''''''''''

Public Property Get Bijzonderheden() As String
    Bijzonderheden = pBijzonderheden
End Property

Public Property Let Bijzonderheden(Value As String)
    pBijzonderheden = Value
End Property

'''''''''''''''''''''''''
' Beoordeling property
'''''''''''''''''''''''''

Public Property Get Beoordeling() As Integer
    Beoordeling = pBeoordeling
End Property

Public Property Let Beoordeling(Value As Integer)
    pBeoordeling = Value
End Property

'''''''''''''''''''''''''
' Actief property
'''''''''''''''''''''''''

Public Property Get Archief() As Boolean
    Archief = pArchief
End Property

Public Property Let Archief(Value As Boolean)
    pArchief = Value
End Property

'''''''''''''''''''''''''
' Bedrijf property
'''''''''''''''''''''''''

Public Property Get BedrijfId() As Long
    BedrijfId = pBedrijfId
End Property

Public Property Let BedrijfId(Value As Long)
    pBedrijfId = Value
End Property

'''''''''''''''''''''''''
' Bedrijf property
'''''''''''''''''''''''''

Public Property Get Bedrijf() As Bedrijf
   Set Bedrijf = pBedrijf
End Property

Public Property Let Bedrijf(Value As Bedrijf)
  Set pBedrijf = Value
End Property


'''''''''''''''''''''''''
' CProducties property
'''''''''''''''''''''''''

Public Property Get CPersoneelPlanning() As Collection
   Set CPersoneelPlanning = pCPersoneelPlanning
End Property

Public Property Let CPersoneelPlanning(Value As Collection)
   Set pCPersoneelPlanning = Value
End Property

Property Get AantalPersoneelPlanningen() As Long
    AantalPersoneelPlanningen = CPersoneelPlanning.Count
End Property

Private Sub class_initialize()
    Set pCPersoneelPlanning = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCPersoneelPlanning = Nothing
End Sub

Public Sub PersoneelPlanningenToevoegen(ByRef pp As PersoneelPlanning)
    pCPersoneelPlanning.Add pp
End Sub


Public Function GetById(Id As Long)
    Dim lijst As Variant
    Dim strSQL As String
    Dim db As New DataBase
   
    strSQL = "SELECT * FROM PERSONEEL where Id = " & Id & ";"
    
    lijst = db.getLijstBySQL(strSQL)
    
    If IsEmpty(lijst) = False Then
        Me.Id = Id
        Me.Achternaam = lijst(1, 0)
        Me.Naam = lijst(2, 0)
        Me.BSN = lijst(3, 0)
        Me.Machinist = lijst(4, 0)
        Me.Timmerman = lijst(5, 0)
        Me.Grondwerker = lijst(6, 0)
        Me.Sloper = lijst(7, 0)
        Me.DHV = lijst(8, 0)
        Me.DTA = lijst(9, 0)
        Me.KVP = lijst(15, 0)
        Me.HVK = lijst(16, 0)
        Me.Uitvoerder = lijst(10, 0)
        Me.Bijzonderheden = lijst(11, 0)
        Me.Beoordeling = lijst(12, 0)
        Me.Archief = lijst(13, 0)
        Me.BedrijfId = lijst(14, 0)
    End If
    
    Set rst = Nothing
    Me.Bedrijf.GetById (Me.BedrijfId)
    
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
    strSQL = "UPDATE [PERSONEEL] " & _
             "SET " & _
             "[Achternaam] = ?, " & _
             "[Naam] = ?, " & _
             "[BSN] = ?, " & _
             "[Machinist] = ?, " & _
             "[Timmerman] = ?, " & _
             "[Grondwerker] = ?, " & _
             "[Sloper] = ?, " & _
             "[DHV] = ?, " & _
             "[DTA] = ?, " & _
             "[Uitvoerder] = ?, " & _
             "[Bijzonderheden] = ?, " & _
             "[Beoordeling] = ?, " & _
             "[Archief] = ?, " & _
             "[BedrijfId] = ?, " & _
             "[KVP] = ?, " & _
             "[HVK] = ? " & _
             "WHERE [Id] = ?;"
    
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Achternaam) + 1, Me.Achternaam)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Naam) + 1, Me.Naam)
        .Parameters.Append .CreateParameter("p3", adBigInt, adParamInput, Len(CStr(Me.BSN)), Me.BSN)
        .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Machinist)), Me.Machinist)
        .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Timmerman)), Me.Timmerman)
        .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.Grondwerker)), Me.Grondwerker)
        .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.Sloper)), Me.Sloper)
        .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.DHV)), Me.DHV)
        .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.DTA)), Me.DTA)
        .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.Uitvoerder)), Me.Uitvoerder)
        .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, Len(CStr(Me.Bijzonderheden)) + 1, Me.Bijzonderheden)
        .Parameters.Append .CreateParameter("p12", adBigInt, adParamInput, Len(CStr(Me.Beoordeling)), Me.Beoordeling)
        .Parameters.Append .CreateParameter("p13", adBoolean, adParamInput, Len(CStr(Me.Archief)), Me.Archief)
        .Parameters.Append .CreateParameter("p14", adBigInt, adParamInput, Len(CStr(Me.BedrijfId)), Me.BedrijfId)
        .Parameters.Append .CreateParameter("p15", adBoolean, adParamInput, Len(CStr(Me.KVP)), Me.KVP)
        .Parameters.Append .CreateParameter("p16", adBoolean, adParamInput, Len(CStr(Me.HVK)), Me.HVK)
        .Parameters.Append .CreateParameter("p17", adBigInt, adParamInput, Len(CStr(Me.Id)), Me.Id)
        
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
    Dim db As New DataBase
    
    'Instert Statement
    strSQL = "INSERT INTO [PERSONEEL] (" & _
                    "[Achternaam]," & _
                    "[Naam]," & _
                    "[BSN]," & _
                    "[Machinist]," & _
                    "[Timmerman]," & _
                    "[Grondwerker]," & _
                    "[Sloper]," & _
                    "[DHV]," & _
                    "[DTA]," & _
                    "[Uitvoerder]," & _
                    "[Bijzonderheden]," & _
                    "[Beoordeling]," & _
                    "[Archief]," & _
                    "[BedrijfId]," & _
                    "[KVP]," & _
                    "[HVK]) " & vbCrLf & _
             "VALUES (p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,p13,p14,p15,p16);"
    
    'Define attachment to database table specifics
    
        db.Connect
        
       
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(CStr(Me.Achternaam)), Me.Achternaam)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Naam) + 1, Me.Naam)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(CStr(Me.BSN)) + 1, Me.BSN)
        .Parameters.Append .CreateParameter("p4", adBoolean, adParamInput, Len(CStr(Me.Machinist)), Me.Machinist)
        .Parameters.Append .CreateParameter("p5", adBoolean, adParamInput, Len(CStr(Me.Timmerman)), Me.Timmerman)
        .Parameters.Append .CreateParameter("p6", adBoolean, adParamInput, Len(CStr(Me.Grondwerker)), Me.Grondwerker)
        .Parameters.Append .CreateParameter("p7", adBoolean, adParamInput, Len(CStr(Me.Sloper)), Me.Sloper)
        .Parameters.Append .CreateParameter("p8", adBoolean, adParamInput, Len(CStr(Me.DHV)), Me.DHV)
        .Parameters.Append .CreateParameter("p9", adBoolean, adParamInput, Len(CStr(Me.DTA)), Me.DTA)
        .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.Uitvoerder)), Me.Uitvoerder)
        .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, Len(CStr(Me.Bijzonderheden)) + 1, Me.Bijzonderheden)
        .Parameters.Append .CreateParameter("p12", adInteger, adParamInput, Len(CStr(Me.Beoordeling)), Me.Beoordeling)
        .Parameters.Append .CreateParameter("p13", adBoolean, adParamInput, Len(CStr(Me.Archief)), Me.Archief)
        .Parameters.Append .CreateParameter("p14", adBigInt, adParamInput, Len(CStr(Me.BedrijfId)), Me.BedrijfId)
        .Parameters.Append .CreateParameter("p15", adBoolean, adParamInput, Len(CStr(Me.KVP)), Me.KVP)
        .Parameters.Append .CreateParameter("p16", adBoolean, adParamInput, Len(CStr(Me.HVK)), Me.HVK)
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
    db.Disconnect
    Set db = Nothing
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    insert = False
    Resume Exit_Insert
End Function


Function ToString() As String
Dim tekst As String

tekst = "PERSONEEL {"
tekst = tekst & vbNewLine & "ID = " & Me.Id
tekst = tekst & vbNewLine & "Achternaam = " & Me.Achternaam
tekst = tekst & vbNewLine & "Naam = " & Me.Naam
tekst = tekst & vbNewLine & "BSN = " & Me.BSN
tekst = tekst & vbNewLine & "Machinist = " & Me.Machinist
tekst = tekst & vbNewLine & "Timmerman = " & Me.Timmerman
tekst = tekst & vbNewLine & "Grondwerker = " & Me.Grondwerker
tekst = tekst & vbNewLine & "Sloper = " & Me.Sloper
tekst = tekst & vbNewLine & "DHV = " & Me.DHV
tekst = tekst & vbNewLine & "DTA = " & Me.DTA
tekst = tekst & vbNewLine & "KVP = " & Me.KVP
tekst = tekst & vbNewLine & "HVK = " & Me.HVK
tekst = tekst & vbNewLine & "Uitvoerder = " & Me.Uitvoerder
tekst = tekst & vbNewLine & "Bijzonderheden = " & Me.Bijzonderheden
tekst = tekst & vbNewLine & "Beoordeling = " & Me.Beoordeling
tekst = tekst & vbNewLine & "Archief = " & Me.Archief
tekst = tekst & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function

