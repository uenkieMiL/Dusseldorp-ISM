VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Materieel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pId As Long
Private pMaterieelcode As String
Private pOmschrijving As String
Private pMerk As String
Private pMaterieelType As String
Private pBouwjaar As Integer
Private pAanschafdatum As Date
Private pKeuringsdatum As Date
Private pSerienummer As String
Private pOnderhoudstermijn As String
Private pLaatsteonderhoudsdatum As Date
Private pFoto As String
Private pInplanbaar As Boolean
Private pInactief As Boolean
Private pfilename As String
Private pfilelocation As String
Private pStatus As String
Private pCMaterieelPlanning As New Collection


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
' MaterieelCode property
'''''''''''''''''''''''''

Public Property Get MaterieelCode() As String
    MaterieelCode = pMaterieelcode
End Property

Public Property Let MaterieelCode(Value As String)
    pMaterieelcode = Value
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
' Merk property
'''''''''''''''''''''''''

Public Property Get Merk() As String
    Merk = pMerk
End Property

Public Property Let Merk(Value As String)
    pMerk = Value
End Property

'''''''''''''''''''''''''
' MaterieelType property
'''''''''''''''''''''''''

Public Property Get MaterieelType() As String
    MaterieelType = pMaterieelType
End Property

Public Property Let MaterieelType(Value As String)
    pMaterieelType = Value
End Property

'''''''''''''''''''''''''
' Bouwjaar property
'''''''''''''''''''''''''

Public Property Get Bouwjaar() As Integer
    Bouwjaar = pBouwjaar
End Property

Public Property Let Bouwjaar(Value As Integer)
    pBouwjaar = Value
End Property

'''''''''''''''''''''''''
' AanschafDatum property
'''''''''''''''''''''''''

Public Property Get AanschafDatum() As Date
    AanschafDatum = pAanschafdatum
End Property

Public Property Let AanschafDatum(Value As Date)
    pAanschafdatum = Value
End Property

'''''''''''''''''''''''''
' KeuringsDatum property
'''''''''''''''''''''''''

Public Property Get KeuringsDatum() As Date
    KeuringsDatum = pKeuringsdatum
End Property

Public Property Let KeuringsDatum(Value As Date)
    pKeuringsdatum = Value
End Property

'''''''''''''''''''''''''
' Serienummer property
'''''''''''''''''''''''''

Public Property Get Serienummer() As String
    Serienummer = pSerienummer
End Property

Public Property Let Serienummer(Value As String)
    pSerienummer = Value
End Property

'''''''''''''''''''''''''
' Onderhoudstermijn property
'''''''''''''''''''''''''

Public Property Get Onderhoudstermijn() As String
    Onderhoudstermijn = pOnderhoudstermijn
End Property

Public Property Let Onderhoudstermijn(Value As String)
    pOnderhoudstermijn = Value
End Property

'''''''''''''''''''''''''
' LaatsteOnderhoudsDatum property
'''''''''''''''''''''''''

Public Property Get LaatsteOnderhoudsDatum() As Date
    LaatsteOnderhoudsDatum = pLaatsteonderhoudsdatum
End Property

Public Property Let LaatsteOnderhoudsDatum(Value As Date)
    pLaatsteonderhoudsdatum = Value
End Property

'''''''''''''''''''''''''
' Foto property
'''''''''''''''''''''''''

Public Property Get Foto() As String
    Foto = pFoto
End Property

Public Property Let Foto(Value As String)
    pFoto = Value
End Property

'''''''''''''''''''''''''
' Inplanbaar property
'''''''''''''''''''''''''

Public Property Get Inplanbaar() As Boolean
    Inplanbaar = pInplanbaar
End Property

Public Property Let Inplanbaar(Value As Boolean)
    pInplanbaar = Value
End Property

'''''''''''''''''''''''''
' InActief property
'''''''''''''''''''''''''

Public Property Get Inactief() As Boolean
    Inactief = pInactief
End Property

Public Property Let Inactief(Value As Boolean)
    pInactief = Value
End Property

'''''''''''''''''''''''''
' Status property
'''''''''''''''''''''''''

Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(Value As String)
    pStatus = Value
End Property


Public Property Get CMaterieelPlanning() As Collection
   Set CMaterieelPlanning = pCMaterieelPlanning
End Property

'Public Property Let CPersoneelPlanning(Value As Collection)
'   Set pCMaterieelPlanning = Value
'End Property

Property Get AantalPersoneelPlanningen() As Long
    AantalPersoneelPlanningen = CMaterieelPlanning.Count
End Property

Private Sub Class_Terminate()
    Set pCMaterieelPlanning = Nothing
End Sub

Public Sub MaterieelPlanningenToevoegen(ByRef mp As MaterieelPlanning)
    pCMaterieelPlanning.Add mp
End Sub


Public Function insert() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim conn As New ADODB.connection
Dim strSQL As String
Dim db As New DataBase
Dim maxid As Variant
    If Me.Status = "" Then Me.Status = "In Magazijn"
    db.Connect
    
    'Instert Statement
    
        strSQL = "INSERT INTO [MATERIEEL] (" & _
                    "[MaterieelCode], " & _
                    "[Omschrijving], " & _
                    "[Merk], " & _
                    "[MaterieelType], " & _
                    "[Bouwjaar], " & _
                    "[AanschafDatum], " & _
                    "[KeuringsDatum], " & _
                    "[Serienummer], " & _
                    "[Onderhoudstermijn], " & _
                    "[LaatsteOnderhoudsDatum], " & _
                    "[Foto], " & _
                    "[Inplanbaar], " & _
                    "[InActief]) " & _
        "VALUES (p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13);"
    
    With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.MaterieelCode) + 1, Me.MaterieelCode)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.Merk) + 1, Me.Merk)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.MaterieelType) + 1, Me.MaterieelType)
        .Parameters.Append .CreateParameter("p5", adInteger, adParamInput, Len(CStr(Me.Bouwjaar)), Me.Bouwjaar)
        .Parameters.Append .CreateParameter("p6", adDate, adParamInput, Len(CStr(Me.AanschafDatum)), Me.AanschafDatum)
        .Parameters.Append .CreateParameter("p7", adDate, adParamInput, Len(CStr(Me.KeuringsDatum)), Me.KeuringsDatum)
        .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, Len(Me.Serienummer) + 1, Me.Serienummer)
        .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, Len(Me.Onderhoudstermijn) + 1, Me.Onderhoudstermijn)
        If Me.LaatsteOnderhoudsDatum <> #12:00:00 AM# Then
        .Parameters.Append .CreateParameter("p10", adDate, adParamInput, Len(CStr(Me.LaatsteOnderhoudsDatum)), Me.LaatsteOnderhoudsDatum)
        Else
        .Parameters.Append .CreateParameter("p10", adDouble, adParamInput, , Null)
        End If
        .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, Len(Me.Foto) + 1, Me.Foto)
        .Parameters.Append .CreateParameter("p12", adBoolean, adParamInput, Len(CStr(Me.Inplanbaar)), Me.Inplanbaar)
        .Parameters.Append .CreateParameter("p13", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
        .CommandText = strSQL
    
        Set rst = .Execute
    End With
    
    'Return a good return code
    insert = True
    
    strSQL = "SELECT MAX([Id]) FROM [MATERIEEL];"
    Set rst = New ADODB.Recordset
    rst.Open strSQL, db.connection, adOpenStatic, adLockReadOnly
    Me.Id = rst.Fields(0).Value
    
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    db.Disconnect
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    insert = False

End Function


Function UpdateFoto()
Dim splitbestand As Variant
Dim splitfilename As Variant

pfilelocation = Application.GetOpenFilename _
(Title:="Selecteer een afbeelding om in te laden", _
FileFilter:="JPG bestand *.jpg* (*.jpg),")
If pfilelocation = "Onwaar" Then Exit Function
If pfilelocation <> "" Then
    splitbestand = Split(pfilelocation, "\")
    pfilename = splitbestand(UBound(splitbestand))
    splitfilename = Split(pfilename, ".")
    
    
    If CheckFolderAangemaakt(ThisWorkbook.Path & "\materieel") = False Then
        MkDir MateriaalLocatie
    End If
    
    If CheckBestandAangemaakt(MateriaalLocatie & Me.Id & "." & splitfilename(UBound(splitfilename))) = True Then
        Kill MateriaalLocatie & Me.Id & "." & splitfilename(UBound(splitfilename))
    End If
    
    FileCopy pfilelocation, MateriaalLocatie & Me.Id & "." & splitfilename(UBound(splitfilename))
    
    Me.Foto = Me.Id & "." & splitfilename(UBound(splitfilename))
    Me.update
End If


End Function


Public Function update() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New DataBase
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [MATERIEEL] " & _
         "SET " & _
         "[MaterieelCode] = ?, " & _
         "[Omschrijving] = ?, " & _
         "[Merk] = ?, " & _
         "[MaterieelType] = ?, " & _
         "[Bouwjaar] = ?, " & _
         "[AanschafDatum] = ?, " & _
         "[KeuringsDatum] = ?, " & _
         "[Serienummer] = ?, " & _
         "[Onderhoudstermijn] = ?, " & _
         "[LaatsteOnderhoudsDatum] = ?, " & _
         "[Foto] = ?, " & _
         "[Inplanbaar] = ?, " & _
         "[InActief] = ?, " & _
         "[Status] = ? " & _
         "WHERE [Id] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL
    .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.MaterieelCode) + 1, Me.MaterieelCode)
    .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Omschrijving) + 1, Me.Omschrijving)
    .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.Merk) + 1, Me.Merk)
    .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.MaterieelType) + 1, Me.MaterieelType)
    .Parameters.Append .CreateParameter("p7", adInteger, adParamInput, Len(CStr(Me.Bouwjaar)), Me.Bouwjaar)
    .Parameters.Append .CreateParameter("p8", adDate, adParamInput, Len(CStr(Me.AanschafDatum)), Me.AanschafDatum)
    .Parameters.Append .CreateParameter("p9", adDate, adParamInput, Len(CStr(Me.KeuringsDatum)), Me.KeuringsDatum)
    .Parameters.Append .CreateParameter("p10", adVarChar, adParamInput, Len(Me.Serienummer) + 1, Me.Serienummer)
    .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, Len(Me.Onderhoudstermijn) + 1, Me.Onderhoudstermijn)
    .Parameters.Append .CreateParameter("p12", adDate, adParamInput, Len(CStr(Me.LaatsteOnderhoudsDatum)), Me.LaatsteOnderhoudsDatum)
    .Parameters.Append .CreateParameter("p13", adVarChar, adParamInput, Len(Me.Foto) + 1, Me.Foto)
    .Parameters.Append .CreateParameter("p14", adBoolean, adParamInput, Len(CStr(Me.Inplanbaar)), Me.Inplanbaar)
    .Parameters.Append .CreateParameter("p15", adBoolean, adParamInput, Len(CStr(Me.Inactief)), Me.Inactief)
    .Parameters.Append .CreateParameter("p16", adVarChar, adParamInput, Len(Me.Status) + 1, Me.Status)
    .Parameters.Append .CreateParameter("p1", adInteger, adParamInput, Len(CStr(Me.Id)), Me.Id)

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
Resume Exit_Insert
End Function


Public Function updateStatus() As Boolean
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim db As New DataBase
Dim con As ADODB.connection
Dim lRecordsAffected As Long

'update Statement
strSQL = "UPDATE [MATERIEEL] " & _
         "SET " & _
         "[Status] = ? " & _
         "WHERE [Id] = ?;"

'Define attachment to database table specifics
db.Connect

With cmd
.ActiveConnection = db.connection
.CommandType = adCmdText
.CommandText = strSQL
    .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Status) + 1, Me.Status)
    .Parameters.Append .CreateParameter("p2", adInteger, adParamInput, Len(CStr(Me.Id)), Me.Id)

Set rst = .Execute(lRecordsAffected)
End With
    
If lRecordsAffected = 0 Then
    updateStatus = False
Else
    'Return a good return code
    updateStatus = True
End If

Exit_Insert:
'Clean up the connection to the database
Set rst = Nothing
Set cmd = Nothing
db.Disconnect

Exit Function
    
Err_Insert:
Call errorhandler_MsgBox("Class: " & TypeName(Me) & " , Function: Insert()")
updateStatus = False
Resume Exit_Insert
End Function


Function FromList(r As Long, lijst As Variant)
pId = lijst(0, r)
If IsNull(lijst(1, r)) = False Then pMaterieelcode = lijst(1, r)
pOmschrijving = lijst(2, r)
pMerk = lijst(3, r)
pMaterieelType = lijst(4, r)
pBouwjaar = lijst(5, r)
If IsNull(lijst(6, r)) = False Or lijst(6, r) = "test" Then pAanschafdatum = lijst(6, r)
If IsNull(lijst(7, r)) = False Then pKeuringsdatum = lijst(7, r)
If IsNull(lijst(8, r)) = False Then pSerienummer = lijst(8, r)
pOnderhoudstermijn = lijst(9, r)
If IsNull(lijst(10, r)) = False Then pLaatsteonderhoudsdatum = lijst(10, r)
If IsNull(lijst(11, r)) = False Then pFoto = lijst(11, r)
pInplanbaar = lijst(12, r)
pInactief = lijst(13, r)
If IsNull(lijst(14, r)) = False Then pStatus = lijst(14, r)
End Function


Function haalop()
Dim db As New DataBase
Dim lijst As Variant

lijst = db.getLijstBySQL("SELECT * FROM MATERIEEL WHERE Id = " & pId & ";")

If IsEmpty(lijst) = False Then Me.FromList 0, lijst
End Function


Function ToString() As String
Dim tekst As String

    tekst = "TEKST {"
    tekst = tekst & vbNewLine & vbTab & "Id = " & Me.Id
    tekst = tekst & vbNewLine & vbTab & "MaterieelCode = " & Me.MaterieelCode
    tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Omschrijving
    tekst = tekst & vbNewLine & vbTab & "Merk = " & Me.Merk
    tekst = tekst & vbNewLine & vbTab & "MaterieelType = " & Me.MaterieelType
    tekst = tekst & vbNewLine & vbTab & "Bouwjaar = " & Me.Bouwjaar
    tekst = tekst & vbNewLine & vbTab & "AanschafDatum = " & Me.AanschafDatum
    tekst = tekst & vbNewLine & vbTab & "KeuringsDatum = " & Me.KeuringsDatum
    tekst = tekst & vbNewLine & vbTab & "Serienummer = " & Me.Serienummer
    tekst = tekst & vbNewLine & vbTab & "Onderhoudstermijn = " & Me.Onderhoudstermijn
    tekst = tekst & vbNewLine & vbTab & "LaatsteOnderhoudsDatum = " & Me.LaatsteOnderhoudsDatum
    tekst = tekst & vbNewLine & vbTab & "Foto = " & Me.Foto
    tekst = tekst & vbNewLine & vbTab & "Inplanbaar = " & Me.Inplanbaar
    tekst = tekst & vbNewLine & vbTab & "InActief = " & Me.Inactief
    tekst = tekst & vbNewLine & vbTab & "Status = " & Me.Status
    tekst = tekst & vbNewLine & "}"
    
    ToString = tekst
End Function

Function Print_r()
    Debug.Print Me.ToString
End Function

