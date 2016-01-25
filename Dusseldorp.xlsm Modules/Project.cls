VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pSynergy As String
Private pOmschrijving As String
Private pOpdrachtgever As String
Private pPV As String
Private pPL As String
Private pCALC As String
Private pWVB As String
Private pUITV As String
Private pNAB As String
Private pOFFERTE As String
Private pVestiging As String
Private pStaatInWacht As Boolean
Private pIntern As String
Private pExtern As String
Private pNaBelDatum As Date
Private pCPlanningen As Collection
Private pCProducties As Collection
Private pCPersoneel As Collection
Private pPlanningVanProject As Planning

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
' Omschrijving property
'''''''''''''''''''''''''

Public Property Get Omschrijving() As String
    Omschrijving = pOmschrijving
End Property

Public Property Let Omschrijving(Value As String)
    pOmschrijving = Value
End Property

'''''''''''''''''''''''''
' Opdrachtgever property
'''''''''''''''''''''''''

Public Property Get Opdrachtgever() As String
    Opdrachtgever = pOpdrachtgever
End Property

Public Property Let Opdrachtgever(Value As String)
    pOpdrachtgever = Value
End Property

'''''''''''''''''''''''''
' PV property
'''''''''''''''''''''''''

Public Property Get pv() As String
    pv = pPV
End Property

Public Property Let pv(Value As String)
    pPV = Value
End Property

'''''''''''''''''''''''''
' PL property
'''''''''''''''''''''''''

Public Property Get pl() As String
    pl = pPL
End Property

Public Property Let pl(Value As String)
    pPL = Value
End Property

'''''''''''''''''''''''''
' CALC property
'''''''''''''''''''''''''

Public Property Get CALC() As String
    CALC = pCALC
End Property

Public Property Let CALC(Value As String)
    pCALC = Value
End Property

'''''''''''''''''''''''''
' WVB property
'''''''''''''''''''''''''

Public Property Get wvb() As String
    wvb = pWVB
End Property

Public Property Let wvb(Value As String)
    pWVB = Value
End Property

'''''''''''''''''''''''''
' UITV property
'''''''''''''''''''''''''

Public Property Get uitv() As String
    uitv = pUITV
End Property

Public Property Let uitv(Value As String)
    pUITV = Value
End Property

'''''''''''''''''''''''''
' NAB property
'''''''''''''''''''''''''

Public Property Get NAB() As String
    NAB = pNAB
End Property

Public Property Let NAB(Value As String)
    pNAB = Value
End Property

'''''''''''''''''''''''''
' OFFERTE property
'''''''''''''''''''''''''

Public Property Get OFFERTE() As String
    OFFERTE = pOFFERTE
End Property

Public Property Let OFFERTE(Value As String)
    pOFFERTE = Value
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
' StaatInWacht property
'''''''''''''''''''''''''

Public Property Get staatInWacht() As Boolean
    staatInWacht = pStaatInWacht
End Property
Public Property Let staatInWacht(Value As Boolean)
    pStaatInWacht = Value
End Property

'''''''''''''''''''''''''
' Intern property
'''''''''''''''''''''''''

Public Property Get intern() As String
    intern = pIntern
End Property
Public Property Let intern(Value As String)
    pIntern = Value
End Property

'''''''''''''''''''''''''
' Extern property
'''''''''''''''''''''''''

Public Property Get extern() As String
    extern = pExtern
End Property
Public Property Let extern(Value As String)
    pExtern = Value
End Property

'''''''''''''''''''''''''
' NaBelDatum property
'''''''''''''''''''''''''

Public Property Get naBelDatum() As Date
    naBelDatum = pNaBelDatum
End Property
Public Property Let naBelDatum(Value As Date)
    pNaBelDatum = Value
End Property

'''''''''''''''''''''''''
' CPlanningen property
'''''''''''''''''''''''''

Public Property Get CPlanningen() As Collection
    Set CPlanningen = pCPlanningen
End Property

Public Property Let CPlanningen(Value As Collection)
    Set pCPlanningen = Value
End Property

'''''''''''''''''''''''''
' CProducties property
'''''''''''''''''''''''''

Public Property Get CProducties() As Collection
   Set CProducties = pCProducties
End Property

Public Property Let CProducties(Value As Collection)
   Set pCProducties = Value
End Property

'''''''''''''''''''''''''
' CProducties property
'''''''''''''''''''''''''

Public Property Get CPersoneel() As Collection
   Set CPersoneel = pCPersoneel
End Property

Public Property Let CPersoneel(Value As Collection)
   Set pCPersoneel = Value
End Property


Public Property Get PlanningVanProject() As Planning
    
     Set PlanningVanProject = pPlanningVanProject
End Property
Public Property Let PlanningVanProject(Value As Planning)
    Set pPlanningVanProject = Value
End Property


Property Get AantalPlaningen() As Long
    AantalPlaningen = pCPlanningen.Count
End Property

Property Get AantalProductie() As Long
    AantalProductie = pCProducties.Count
End Property

Property Get AantalPersoneel() As Long
    AantalPersoneel = pCPersoneel.Count
End Property

Property Get PlanningItem(ByVal vID As Variant) As Planning
    Set PlanningItem = pCPlanningen(vID)
End Property

Property Get ProductieItem(ByVal vID As Variant) As project
    Set ProductieItem = pCProducties(vID)
End Property

Property Get PersoneelItem(ByVal vID As Variant) As Personeel
    Set PersoneelItem = PersoneelItem(vID)
End Property

Private Sub class_initialize()
    Set pCPlanningen = New Collection
    Set pCProducties = New Collection
    Set pPlanningVanProject = New Planning
    Set pCPersoneel = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCPlanningen = Nothing
    Set pCProducties = Nothing
    Set pPlanningVanProject = Nothing
    Set pCPersoneel = Nothing
End Sub

Public Sub ToevoegenPlanningen(ByRef p As Planning)
    Dim ap As Planning
    Set ap = New Planning
    Set ap = p
    pCPlanningen.Add ap
End Sub

Public Sub ToevoegenProductie(ByRef p As Productie)
    Dim ap As Productie
    Set ap = New Productie
    Set ap = p
    pCProducties.Add ap
End Sub

Public Sub ToevoegenPersoneel(ByRef p As Personeel)
    pCPersoneel.Add p
End Sub

Public Function update() As Boolean
    Dim strSQL As String
    Dim db As New DataBase
    Dim lRecordsAffected As Long
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    'update Statement
    strSQL = "UPDATE [PROJECTEN] " & _
             "SET " & _
             "[Omschrijving] = ?," & _
             "[Opdrachtgever] = ?," & _
             "[PV] = ?," & _
             "[PL] = ?," & _
             "[CALC] = ?," & _
             "[WVB] = ?," & _
             "[UITV] = ?," & _
             "[NAB] = ?," & _
             "[OFFERTE] = ?," & _
             "[WACHT] = ? " & _
             "WHERE [Synergy] = ? " & _
             "AND [Vestiging] = ?;"
    
    'Define attachment to database table specifics
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adVarChar, adParamInput, Len(Me.Omschrijving), Me.Omschrijving)
        .Parameters.Append .CreateParameter("p2", adVarChar, adParamInput, Len(Me.Opdrachtgever), Me.Opdrachtgever)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.pv), Me.pv)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.pl), Me.pl)
        .Parameters.Append .CreateParameter("p5", adVarChar, adParamInput, Len(Me.CALC), Me.CALC)
        .Parameters.Append .CreateParameter("p6", adVarChar, adParamInput, Len(Me.wvb), Me.wvb)
        .Parameters.Append .CreateParameter("p7", adVarChar, adParamInput, Len(Me.uitv), Me.uitv)
        .Parameters.Append .CreateParameter("p8", adVarChar, adParamInput, Len(Me.NAB) + 1, Me.NAB)
        .Parameters.Append .CreateParameter("p9", adVarChar, adParamInput, Len(Me.OFFERTE), Me.OFFERTE)
        .Parameters.Append .CreateParameter("p10", adBoolean, adParamInput, Len(CStr(Me.staatInWacht)), Me.staatInWacht)
        .Parameters.Append .CreateParameter("p11", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p12", adVarChar, adParamInput, Len(Me.Vestiging), Me.Vestiging)
        
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

Function haalop()
Dim planningid As Long
Dim strSQL As String
Dim lijstprojecten As Variant
Dim lijstPlanningen As Variant
Dim lijstProductie As Variant
Dim pl As Planning
Dim pr As Productie
Dim t As taak
Dim x As Long
Dim lt As Long
Dim lpl As Long

Dim db As New DataBase

strSQL = "SELECT * FROM PROJECTEN WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "';"

lijstprojecten = db.getLijstBySQL(strSQL)

strSQL = "SELECT * FROM PLANNINGEN WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "';"

lijstPlanningen = db.getLijstBySQL(strSQL)
If IsEmpty(lijstPlanningen) = False Then planningid = lijstPlanningen(0, 0)

strSQL = "SELECT * FROM PRODUCTIE INNER JOIN  PRODUCTIESOORT ON PRODUCTIE.Soort = PRODUCTIESOORT.Id WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "';"
lijstProductie = db.getLijstBySQL(strSQL)

If planningid > 0 Then
    strSQL = "SELECT * FROM TAKEN WHERE PlanningId = " & planningid & ";"
    lijstTaken = db.getLijstBySQL(strSQL)
End If
If IsEmpty(lijstprojecten) = False Then
    For x = 0 To UBound(lijstprojecten, 2)
        FromList x, lijstprojecten
        
        For lpl = 0 To UBound(lijstPlanningen, 2)
            Set pl = New Planning
            pl.FromList lpl, lijstPlanningen
            
            If IsEmpty(lijstTaken) = False Then
                For lt = 0 To UBound(lijstTaken, 2)
                    
                    Set t = New taak
                    t.FromList lt, lijstTaken
                    pl.ToevoegenTaak t
                Next lt
            End If
            Me.ToevoegenPlanningen pl
            
        Next lpl
        
        If IsEmpty(lijstProductie) = False Then
            For lpr = 0 To UBound(lijstProductie, 2)
                Set pr = New Productie
                pr.Id = lijstProductie(0, lpr)
                pr.soort = lijstProductie(3, lpr)
                pr.startdatum = lijstProductie(4, lpr)
                pr.einddatum = lijstProductie(5, lpr)
                pr.Kleur = lijstProductie(10, lpr)
                pr.Gereed = lijstProductie(7, lpr)
                pr.Omschrijving = lijstProductie(9, lpr)
                
                Me.ToevoegenProductie pr
            Next lpr
        End If
    Next x
End If

End Function


Function ToString() As String
Dim tekst As String

tekst = "PROJECT " & vbNewLine & "{"

tekst = tekst & vbNewLine & vbTab & "Synergy = " & Me.synergy
tekst = tekst & vbNewLine & vbTab & "PlanningId = " & Me.Omschrijving
tekst = tekst & vbNewLine & vbTab & "Omschrijving = " & Me.Opdrachtgever
tekst = tekst & vbNewLine & vbTab & "Uitvoeren = " & Me.pv
tekst = tekst & vbNewLine & vbTab & "Volgnummer  = " & Me.pl
tekst = tekst & vbNewLine & vbTab & "Startdatum = " & Me.CALC
tekst = tekst & vbNewLine & vbTab & "Einddatum = " & Me.wvb
tekst = tekst & vbNewLine & vbTab & "Aantal = " & Me.uitv
tekst = tekst & vbNewLine & vbTab & "Ehd = " & Me.Vestiging
tekst = tekst & vbNewLine & vbTab & "Intern = " & Me.intern
tekst = tekst & vbNewLine & vbTab & "Extern = " & Me.extern
tekst = tekst & vbNewLine & "}"

ToString = tekst
End Function

Function Print_r()
Debug.Print ToString
End Function

Function verwijderenProject()
Dim pl As New Planning
Dim Begrotingen As Long
Dim begrotingregels As Long
Dim behoeften As Long
Dim Beschikbaarheid As Long
Dim personeelplanningsregels As Long
Dim planningen As Long
Dim productieperioden As Long
Dim taken As Long
Dim project As Long
Dim tekst As String
Dim ap As Long
Dim lijst As Variant
Dim db As New DataBase
lijst = db.getLijstBySQL("SELECT COUNT(*) FROM PROJECTEN WHERE Synergy = '" & Me.synergy & "';")

If IsEmpty(lijst) = True And lijst(0, 0) = 0 Then Exit Function

For Each pl In Me.CPlanningen
    taken = taken + db.UpdateQueryUitvoeren("DELETE FROM TAKEN WHERE PlanningId = " & pl.Id)
Next pl
    Begrotingen = db.UpdateQueryUitvoeren("DELETE FROM BEGROTINGEN WHERE Synergy = '" & Me.synergy & "';")
    begrotingregels = db.UpdateQueryUitvoeren("DELETE FROM BEGROTINGREGELS WHERE Synergy = '" & Me.synergy & "';")
    behoeften = db.UpdateQueryUitvoeren("DELETE FROM BEHOEFTEN WHERE Synergy = '" & Me.synergy & "';")
    If lijst(0, 0) = 1 Then
    personeelplanningsregels = db.UpdateQueryUitvoeren("DELETE FROM PLANNING_PERSONEEL WHERE Synergy = '" & Me.synergy & "';")
    End If
    planningen = db.UpdateQueryUitvoeren("DELETE FROM PLANNINGEN WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "'; ")
    productieperioden = db.UpdateQueryUitvoeren("DELETE FROM PRODUCTIE WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "';")
    project = db.UpdateQueryUitvoeren("DELETE FROM PROJECTEN WHERE Synergy = '" & Me.synergy & "' AND Vestiging = '" & Me.Vestiging & "';")

tekst = "Het project met Synergy nummer " & Me.synergy & " is succesvol verwijderd. De volgende zaken zijn in de database verwijderd"
tekst = tekst & vbNewLine & Begrotingen & " Begroting(en) zijn er verwijderd"
tekst = tekst & vbNewLine & begrotingregels & " begrotingregel(s) zijn er verwijderd"
tekst = tekst & vbNewLine & behoeften & " Behoefte(n) zijn er verwijderd"
tekst = tekst & vbNewLine & personeelplanningsregels & " regel(s) in de planning zijn er verwijderd"
tekst = tekst & vbNewLine & planningen & " planning(en) zijn er verwijderd"
tekst = tekst & vbNewLine & productieperioden & " productie(s) zijn er verwijderd"
If project = 1 Then tekst = tekst & vbNewLine & "Project is succesvol uit de database verwijderd." Else tekst = tekst & vbNewLine & "Project is NIET uit de database verwijderd."

MsgBox tekst, vbInformation, "VERWIJDEREN PROJECT"
End Function

Function FromList(r As Long, lijst As Variant)
    Me.synergy = lijst(0, r)
    Me.Omschrijving = lijst(1, r)
    Me.Opdrachtgever = lijst(2, r)
    Me.pv = lijst(3, r)
    Me.pl = lijst(4, r)
    Me.CALC = lijst(5, r)
    Me.wvb = lijst(6, r)
    Me.uitv = lijst(7, r)
    If IsNull(lijst(8, r)) = False Then Me.NAB = lijst(8, r)
    Me.OFFERTE = lijst(9, r)
    Me.Vestiging = lijst(10, r)
    Me.staatInWacht = lijst(12, r)
    If IsNull(lijst(13, r)) = False Then Me.naBelDatum = lijst(13, r)
    If IsNull(lijst(14, r)) = False Then Me.intern = lijst(14, r)
    If IsNull(lijst(15, r)) = False Then Me.intern = lijst(15, r)
End Function

Function UpdateWacht() As Boolean

 Dim strSQL As String
    Dim db As New DataBase
    Dim lRecordsAffected As Long
    Dim cmd As New ADODB.Command
    Dim rst As New ADODB.Recordset
    'update Statement
    strSQL = "UPDATE [PROJECTEN] " & _
             "SET " & _
             "[WACHT] = ?," & _
             "[NABELLEN] = ? " & _
             "WHERE [Synergy] = ? " & _
             "AND [Vestiging] = ?;"
    
    'Define attachment to database table specifics
        db.Connect
        
        With cmd
        .ActiveConnection = db.connection
        .CommandType = adCmdText
        .CommandText = strSQL
        .Parameters.Append .CreateParameter("p1", adBoolean, adParamInput, Len(CStr(Me.staatInWacht)), Me.staatInWacht)
        .Parameters.Append .CreateParameter("p2", adDate, adParamInput, Len(CStr(Me.naBelDatum)), Me.naBelDatum)
        .Parameters.Append .CreateParameter("p3", adVarChar, adParamInput, Len(Me.synergy), Me.synergy)
        .Parameters.Append .CreateParameter("p4", adVarChar, adParamInput, Len(Me.Vestiging), Me.Vestiging)
        
        Set rst = .Execute(lRecordsAffected)
    End With
    
    If lRecordsAffected = 0 Then
      UpdateWacht = False
    Else
      'Return a good return code
      UpdateWacht = True
    End If
    
Exit_Insert:
    'Clean up the connection to the database
    Set rst = Nothing
    Set cmd = Nothing
    db.Disconnect
    
    Exit Function
    
Err_Insert:
    Call errorhandler_MsgBox("Class: " & TypeName(Me) & ", Function: Insert()")
    UpdateWacht = False
    Resume Exit_Insert

'l.createLog "Nabeldatum verander van " & CStr(datumoud) & " in " & CStr(datum), pr_updaten, synergy, project

End Function

