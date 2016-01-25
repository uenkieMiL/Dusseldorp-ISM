Attribute VB_Name = "Functies"
Private screenUpdateState As Boolean
Private statusBarState As Boolean
Private calcState As Long
Private eventsState As Boolean
Private displayPageBreakState As Boolean

Public Const locatieprojecten = "J:\Allen\Projecten\"

Sub Turbo_AAN()
'Get current state of various Excel settings; put this at the beginning of your code
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting
End Sub

Sub turbo_UIT()
'after your code runs, restore state; put this at the end of your code
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.Calculation = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True 'note this is a sheet-level setting
End Sub

Sub errorhandler_MsgBox(error As String)
MsgBox error, vbCritical, "FOUT"
End Sub

Function Kolombreedte(breedte As Long, rng As Range, ws As String)
Dim str As Variant

str = Split(rng.Address, "$")
With ThisWorkbook.Worksheets(ws).Range(str(1) & ":" & str(3))
.ColumnWidth = breedte

End With
End Function

Function TaakBalkPlaatsen(CKalender As Collection, startdatum As Date, einddatum As Date, Status As Boolean, r As Long, uitvoeren As Boolean)
        Dim d As datum
        Dim k1 As Long: k1 = -1
        Dim k2 As Long: k2 = -1
        Dim Kleur As Long
        Dim lk As Long
        Dim rng As Range
        ws = ActiveSheet.Name
        If Status = True Then Kleur = 5287936 Else Kleur = 192
        On Error Resume Next
            Set d = CKalender.item(CStr(startdatum))
            If d.datum = "0:00:00" Then k1 = -1 Else k1 = k.Kolomnummer
            Set d = New datum
            Set d = CKalender.item(CStr(einddatum))
            If k.datum = "0:00:00" Then k2 = -1 Else k2 = k.Kolomnummer
        Resume Next
        If k2 = -1 And k1 = -1 Then Exit Function
        If einddatum = #12:00:00 AM# Then uitvoeren = False: k1 = 0: k2 = lk
        Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, sk + k1), Cells(r, sk + k2))
        If uitvoeren = True Then rng.Interior.Color = Kleur Else rng.Interior.Color = xlNone
        
End Function

Function ProductieBalkPlaatsen(CKalender As Collection, startdatum As Date, einddatum As Date, Kleur As Long, r As Long, ws As String)
        Dim d As datum
        Dim k1 As Long
        Dim k2 As Long
        Dim lk As Long
        Dim rng As Range
        ws = ActiveSheet.Name
        For Each k In CKalender
        If startdatum = k.datum Then k1 = k.Kolomnummer
        If einddatum = k.datum Then k2 = k.Kolomnummer
        If k1 <> 0 And k2 <> 0 Then Exit For
        lk = k.Kolomnummer
        Next k
        If einddatum = #12:00:00 AM# Then uitvoeren = False: k1 = 0: k2 = lk
        Set rng = ThisWorkbook.Sheets(ws).Range(Cells(r, sk + k1), Cells(r, sk + k2))
        rng.Interior.Color = Kleur
        
End Function

Function MaakLijstJN(rng As Range)
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="J,N"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function MaakLijstIenE(rng As Range)
    With rng.Validation
        .delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="JA,NEE,NVT"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function PlaatsFeestdagen(CKalender As Collection, ws As String, sk As Integer)
Dim d As datum
For Each d In CKalender
        If d.Kolomnummer > -1 And d.feestdag = True Then
            Set rng = ThisWorkbook.Sheets(ws).Range(Cells(5, sk + d.Kolomnummer), Cells(1048576, sk + d.Kolomnummer))
            rng.Interior.Color = 12566463
        End If
    Next d
End Function
Function MaakKolomVandaag(CKalender As Collection, ws As String, sk As Integer)
Dim vandaag As Date: vandaag = Now(): vandaag = FormatDateTime(Date, vbShortDate)
Dim d As datum
Dim gevonden As Boolean
Dim kolom As Long
Dim rng As Range

Do While gevonden = False
    For Each d In CKalender
        If d.datum = vandaag Then
            kolom = d.Kolomnummer
            gevonden = True
            Exit For
        End If
    Next d
    
    If gevonden = False Then vandaag = DateAdd("d", 1, vandaag)
Loop

Sheets(ws).Select
Set rng = ThisWorkbook.Sheets(ws).Range(Cells(5, sk + kolom), Cells(1048576, sk + kolom))
rng.Interior.Color = kleuren.vandaag
End Function

Function SelecteerKolomVandaag(CKalender As Collection, ws As String, sk As Integer)
Dim vandaag As Date: vandaag = Now(): vandaag = FormatDateTime(Date, vbShortDate)
Dim d As datum
Dim gevonden As Boolean
Dim kolom As Long
Dim rng As Range

Do While gevonden = False
    For Each d In CKalender
        If d.datum = vandaag Then
            kolom = d.Kolomnummer
            gevonden = True
            Exit For
        End If
    Next d
    
    If gevonden = False Then vandaag = DateAdd("d", 1, vandaag)
Loop

ThisWorkbook.Sheets(ws).Range(Cells(ActiveCell.Row, sk + kolom).Address).Select
End Function



Function DagJaarSamenvoegen(ws As String, rij As Long)
Dim lk As Long
Dim sk1 As Long
Dim k As Long
Dim rng As Range
Dim dezekolom As Long, vorigekolom As Long


If ws = "" Then Exit Function

sk1 = SoortPlanning.startkolom
lk = ThisWorkbook.Sheets(ws).Range("XFD" & rij).End(xlToLeft).Column
Application.DisplayAlerts = False
For k = SoortPlanning.startkolom + 1 To lk + 1
    dezekolom = ThisWorkbook.Sheets(ws).Range(Cells(rij, k).Address)
    vorigekolom = ThisWorkbook.Sheets(ws).Range(Cells(rij, k - 1).Address)
    If dezekolom <> vorigekolom Then
    Set rng = ThisWorkbook.Sheets(ws).Range(Range(Cells(rij, sk1), Cells(rij, k - 1)).Address)
    With rng
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
        With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    sk1 = k
End If

Next k
Application.DisplayAlerts = True
End Function

Function DagMaandSamenvoegen(ws As String, rij As Long)
Dim lk As Long
Dim sk1 As Long
Dim k As Long
Dim rng As Range

sk1 = SoortPlanning.startkolom
lk = ThisWorkbook.Sheets(ws).Range("XFD2").End(xlToLeft).Column
Application.DisplayAlerts = False
For k = sk1 + 1 To lk + 1
    If ThisWorkbook.Sheets(ws).Range(Cells(rij, k).Address) <> ThisWorkbook.Sheets(ws).Range(Cells(rij, k - 1).Address) Then

    Set rng = ThisWorkbook.Sheets(ws).Range(Range(Cells(rij, sk1), Cells(rij, k - 1)).Address)
    With rng
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
        With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With

        
            With ThisWorkbook.Sheets(ws).Range(Range(Cells(rij, sk1), Cells(rij, k - 1)).Address)
        .Merge
        .HorizontalAlignment = xlCenter
    End With


    sk1 = k
End If

Next k
Application.DisplayAlerts = True
End Function


Function DagWeekSamenvoegen(ws As String, rij As Long)
Dim lk As Long
Dim sk1 As Long
Dim k As Long
Dim rng As Range

sk1 = SoortPlanning.startkolom
lk = ThisWorkbook.Sheets(ws).Range("XFD3").End(xlToLeft).Column
Application.DisplayAlerts = False
For k = sk1 + 1 To lk + 1
    If ThisWorkbook.Sheets(ws).Range(Cells(rij, k).Address) <> ThisWorkbook.Sheets(ws).Range(Cells(rij, k - 1).Address) Then

    
        With ThisWorkbook.Sheets(ws).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
        End With
    
        With ThisWorkbook.Sheets(ws).Range(Range(Columns(sk1), Columns(k - 1)).Address).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
        End With
        
        Set rng = ThisWorkbook.Sheets(ws).Range(Range(Cells(rij, sk1), Cells(rij, k - 1)).Address)
        With rng
            .Merge
            .HorizontalAlignment = xlCenter
        End With
    
        With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
        
    With ThisWorkbook.Sheets(ws).Range(Range(Cells(rij, sk1), Cells(rij, k - 1)).Address)
        .Merge
        .HorizontalAlignment = xlCenter
    End With

    
    sk1 = k
  End If
Next k
  
End Function

Function DikkeStrepen(ws As String, CKalendar As Collection, sk As Long)
    Dim w As Long
    Dim d As datum

        For Each d In CKalendar
        If w <> IsoWeekNumber(d.datum) Then
            With ThisWorkbook.Sheets(ws).Range(Columns(sk + d.Kolomnummer).Address).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            End If
         w = IsoWeekNumber(d.datum)
        Next d
   
End Function
Public Function IsoWeekNumber(d1 As Date) As Integer
   'Attributed to Daniel Maher
   Dim d2 As Long
   d2 = DateSerial(Year(d1 - Weekday(d1 - 1) + 4), 1, 3)
   IsoWeekNumber = Int((d1 - d2 + Weekday(d2) + 5) / 7)
End Function

Function MaakRaster(rng As Range)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Function

Function SoortnaarString(soort As Byte) As String
    Select Case soort
        Case 1
        SoortnaarString = "Acquisitie"
        
        Case 2
        SoortnaarString = "Calculatie"
        
        Case 3
        SoortnaarString = "Werkvoorbereiding"
        
        Case 4
        SoortnaarString = "Uitvoering"
    End Select
End Function

Function LijstOpBasisVanQuery(querynaam As String) As Variant

Dim db As New DataBase

LijstOpBasisVanQuery = db.getLijstBySQL(querynaam)

End Function

Function GeefCellectieProductiesVoorProject(ByRef lijst As Collection, synergy As String, Vestiging As String) As Collection
    Dim pr As Productie
    Set GeefCellectieProductiesVoorProject = New Collection
    
    For Each pr In lijst
        If pr.synergy = synergy And pr.Vestiging = Vestiging Then
            GeefCellectieProductiesVoorProject.Add pr
        End If
    Next pr
    
    For Each pr In GeefCellectieProductiesVoorProject
        lijst.Remove CStr(pr.Id)
    
    Next pr
End Function

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded



Function ZoekNieuweDatumVoorTaak(datum As Date, dagen As Long, CKalender As Collection) As Date
Dim d As datum
Dim k As Long

For Each d In CKalender
If d.datum = datum Then k = d.Kolomnummer: Exit For
Next d

k = k + dagen

For Each d In CKalender
If d.Kolomnummer = k Then ZoekNieuweDatumVoorTaak = d.datum: Exit For
Next d

End Function

Function CheckProjectIsAangemaakt(synergy As String, Vestiging As String) As Boolean
Dim lijst As Variant
Dim db As New DataBase
    lijst = db.getLijstBySQL("SELECT Count(*) AS AANTAL FROM PROJECTEN WHERE (((PROJECTEN.Synergy)='" & synergy & "' AND Vestiging = '" & Vestiging & "'));")
    If lijst(0, 0) > 0 Then CheckProjectIsAangemaakt = True
End Function

Function KolomNaarDatumviaNummer(k As Long, CKalender As Collection) As Date
    Dim d As New datum
    
    For Each d In CKalender
        If d.Kolomnummer = k - SoortPlanning.startkolom Then
            KolomNaarDatumviaNummer = d.datum
            Exit For
        End If
    Next d
    
End Function

Function TaakBalkPlaatsenMulti(k1 As Long, k2 As Long, Status As Boolean, rij As Long, uitvoeren As Boolean)
        Dim Kleur As Long
        Dim rng As Range
        Dim ws As String
        ws = ActiveSheet.Name

        If Status = True Then Kleur = kleuren.taak_gereed Else Kleur = kleuren.taak_niet_gereed
        Set rng = ThisWorkbook.Sheets(ws).Range(Cells(rij, k1), Cells(rij, k2))
        If uitvoeren = True Then rng.Interior.Color = Kleur Else rng.Interior.Color = xlNone
        
End Function

Function verwijderoudebalk(rij As Long, t As taak)
Dim rng As Range
Dim ws As String
ws = ActiveSheet.Name
Dim lk As Long: lk = ThisWorkbook.Sheets(ws).Range("XFD" & SoortPlanning.startrij - 1).End(xlToLeft).Column

Set rng = ThisWorkbook.Sheets(ws).Range(Cells(rij, SoortPlanning.startkolom), Cells(rij, lk))
With rng
    .Interior.Color = xlNone
End With

End Function

Function DatumAantalWerkDagenVerplaatsenCollection(oudedatum As Date, aantaldagen As Integer, CKalender As Collection)
Dim d As datum
Dim kolom As Long

Set d = CKalender.item(CStr(oudedatum))
kolom = d.Kolomnummer + aantaldagen
For Each d In CKalender
    If d.Kolomnummer = kolom Then
        Exit For
    End If
Next d

DatumAantalWerkDagenVerplaatsenCollection = d.datum

End Function



Function OpenFolder(synergy As String)
Dim bestand As String, map As String, textline As String, Locatie As String
bestand = "get_dirs.txt"
bestand = ThisWorkbook.Path & "\" & bestand
If Not (Dir$(bestand) <> "") Then Exit Function
Open bestand For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If Left(textline, 5) = synergy Then
    map = textline
    Exit Do
    End If
Loop

Close #1
Locatie = locatieprojecten & map & "\"
Shell "EXPLORER.EXE " & Chr(34) & Locatie & Chr(34), vbNormalFocus
End Function

Function GetLocatie(synergy As String) As String
Dim bestand As String, map As String, textline As String, Locatie As String
bestand = "get_dirs.txt"
bestand = ThisWorkbook.Path & "\" & bestand
If (Dir$(bestand) <> "") Then
Open bestand For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If Left(textline, 5) = synergy Then
    map = textline
    Exit Do
    End If
Loop

Close #1
    If map <> "" Then
        GetLocatie = locatieprojecten & map & "\"
        Exit Function
    End If
End If

End Function

Public Function KolomNaarLetter(Column As Long) As String
    If Column < 1 Then Exit Function
    KolomNaarLetter = KolomNaarLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function

Function Maandnaam(nr As Long) As String
    Select Case nr
        Case 1
        Maandnaam = "Januari"
        Case 2
        Maandnaam = "Februari"
        Case 3
        Maandnaam = "Maart"
        Case 4
        Maandnaam = "April"
        Case 5
        Maandnaam = "Mei"
        Case 6
        Maandnaam = "Juni"
        Case 7
        Maandnaam = "Juli"
        Case 8
        Maandnaam = "Augustus"
        Case 9
        Maandnaam = "September"
        Case 10
        Maandnaam = "Oktober"
        Case 11
        Maandnaam = "November"
        Case 12
        Maandnaam = "December"
    End Select
End Function


Function CheckFolderAangemaakt(pad As String) As Boolean

    Dim FSO As Scripting.FileSystemObject
    Dim FolderPath As String

    Set FSO = New Scripting.FileSystemObject

    FolderPath = pad
    If Right(FolderPath, 1) <> "\" Then
        FolderPath = FolderPath & "\"
    End If

    If FSO.FolderExists(FolderPath) = False Then
        CheckFolderAangemaakt = False
    Else
        CheckFolderAangemaakt = True
    End If

End Function

Function CheckBestandAangemaakt(bestand As String) As Boolean

    Dim FSO As Scripting.FileSystemObject
    
    Set FSO = New Scripting.FileSystemObject

    If FSO.FileExists(bestand) = False Then
        CheckBestandAangemaakt = False
    Else
        CheckBestandAangemaakt = True
    End If

End Function

Function MateriaalLocatie() As String
MateriaalLocatie = ThisWorkbook.Path & "\materieel\"
End Function

Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = col.item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

Function SelectRegel() As Long
Dim tekst As String
Dim s1 As Variant
Dim s2 As Variant
    tekst = Application.InputBox(prompt:="Selecteer een regel", Type:=0)

    'controleer of er wel een range is geselecteerd

    If tekst = "Onwaar" Then Exit Function
    
    If InStr(1, tekst, ":", vbTextCompare) Then
        s1 = Split(tekst, ":")
        tekst = s1(0)
    End If
        tekst = Right(tekst, Len(tekst) - 2)
        s2 = Split(tekst, "C")
        SelectRegel = CLng(s2(0))
End Function

Function DatumNaarKolomnummer(CKalender As Collection, datum As Date) As Long
Dim d As datum
Dim d1 As datum
DatumNaarKolomnummer = -1
If Functies.InCollection(CKalender, CStr(datum)) = True Then
    Set d = CKalender(CStr(datum))
    If d.Kolomnummer > -1 Then DatumNaarKolomnummer = d.Kolomnummer
    Do While DatumNaarKolomnummer = -1
     datum = DateAdd("d", -1, datum)
         If Functies.InCollection(CKalender, CStr(datum)) = True Then
             Set d1 = CKalender(CStr(datum))
             If d1.Kolomnummer > -1 Then DatumNaarKolomnummer = d1.Kolomnummer
         End If
    Loop
       
End If

End Function
