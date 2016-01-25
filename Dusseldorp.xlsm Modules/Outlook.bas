Attribute VB_Name = "Outlook"
Function AgendaItemToevoegen()
'Dim l As New Log
Dim t As New taak
Dim p As New project
Dim synergy As String
Dim time2 As String
Dim time1 As String
Dim ws As String
Dim starttijd As Variant
Dim eindtijd As Variant
Dim Vestiging As String

ws = ActiveSheet.Name
synergy = ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)
Vestiging = ThisWorkbook.Sheets(ws).Range("B" & ActiveCell.Row)
p.synergy = synergy
p.Vestiging = Vestiging
p.haalop
t.haalop (ThisWorkbook.Sheets(ws).Range("V" & ActiveCell.Row))

         
' On Error Resume Next
Dim tekst As String

tekst = "Synergy:" & vbNewLine & p.synergy & vbNewLine & vbNewLine & _
                 "Project Omschrijving:" & vbNewLine & p.Omschrijving & vbNewLine & vbNewLine & _
                 "Opdrachtgever:" & vbNewLine & p.Opdrachtgever & vbNewLine & vbNewLine & _
                 "Taak Omschrijving:" & vbNewLine & t.Omschrijving & vbNewLine & vbNewLine & _
                 "Datum in agenda:" & vbNewLine & t.startdatum & vbNewLine & vbNewLine
time1 = InputBox(tekst & "Geef de startijd op" & vbNewLine & "opmaak = UU:MM", "STARTIJD OPGEVEN")

If time1 = "" Then Exit Function
time1 = Replace(time1, ";", ":")
time1 = Replace(time1, ",", ":")
time1 = Replace(time1, ".", ":")
If IsDate(time1) = False Then
    MsgBox "Geen geldige tijd, probeer opnieuw", vbCritical, "FOUTIEVE WAARDE"
    Exit Function
End If

starttijd = Split(time1, ":", -1, vbBinaryCompare)

time2 = InputBox(tekst & "Geef de eindtijd op" & vbNewLine & "opmaak = UU:MM", "EINDTIJD OPGEVEN")
If time2 = "" Then Exit Function
time2 = Replace(time2, ";", ":")
time2 = Replace(time2, ",", ":")
time2 = Replace(time2, ".", ":")
If IsDate(time2) = False Then
    MsgBox "Geen geldige tijd, probeer opnieuw", vbCritical, "FOUTIEVE WAARDE"
    Exit Function
End If


eindtijd = Split(time2, ":", -1, vbBinaryCompare)
t.startdatum = t.startdatum + TimeValue(starttijd(0) & ":" & starttijd(1) & ":00")
t.einddatum = t.einddatum + TimeValue(eindtijd(0) & ":" & eindtijd(1) & ":00")

'Resume Next
    With CreateObject("Outlook.Application").CreateItem(1)
        .Start = t.startdatum
        .Duration = DateDiff("n", t.startdatum, t.einddatum)
        .Subject = t.Omschrijving & " - " & p.synergy & " - " & p.Omschrijving & " - " & p.Opdrachtgever
        '.Location = "Aanvullen"
        .Body = t.Opmerking
        .save
        
          'l.createLog "Aanmaken TaakItem Outlook, " & p.synergy & " - " & p.omschrijving & " - " & t.omschrijving & ", datum = " & t.startdatum, overzicht_gemaakt, "OUTLOOK AGENDA", 5
    End With
End Function
Function TaakItemToevoegen()
Dim t As New taak
Dim p As New project
Dim synergy As String
Dim time2 As Byte
Dim time1 As Byte
'Dim l As New Log
Dim ws As String
Dim Vestiging As String

ws = ActiveSheet.Name
synergy = ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)
Vestiging = ThisWorkbook.Sheets(ws).Range("B" & ActiveCell.Row)
p.synergy = synergy
p.Vestiging = Vestiging
p.haalop
t.haalop (ThisWorkbook.Sheets(ws).Range("V" & ActiveCell.Row))

tekst = "Synergy:" & vbNewLine & p.synergy & vbNewLine & vbNewLine & _
                 "Project Omschrijving:" & vbNewLine & p.Omschrijving & vbNewLine & vbNewLine & _
                 "Opdrachtgever:" & vbNewLine & p.Opdrachtgever & vbNewLine & vbNewLine & _
                 "Taak Omschrijving:" & vbNewLine & t.Omschrijving & vbNewLine & vbNewLine & _
                 "Vervaldatum van de taak:" & vbNewLine & t.startdatum & vbNewLine & vbNewLine & _
                 "Weet u zeker dat u deze taak wilt aanmaken?"
                
antwoord = MsgBox(tekst, vbYesNo, "OUTLOOK TAAK AANMAKEN")

If antwoord = vbYes Then
    With CreateObject("Outlook.Application").CreateItem(3)
           .Subject = p.synergy & " - " & p.Omschrijving & " - " & t.Omschrijving
           .StartDate = t.startdatum
           .DueDate = t.einddatum
           .ReminderTime = .StartDate - 1
           .ReminderSet = True
           .Body = t.Omschrijving & vbNewLine & t.Opmerking
           .save
    End With
    'l.createLog "Aanmaken TaakItem Outlook, " & p.synergy & " - " & p.omschrijving & " - " & t.omschrijving & ", datum = " & t.startdatum, overzicht_gemaakt, "OUTLOOK TAAK", 5
End If
    
End Function

Sub Mail_ActiveSheet()
Attribute Mail_ActiveSheet.VB_ProcData.VB_Invoke_Func = "m\n14"
'Working in Excel 2000-2013
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = ActiveWorkbook
    
    'Copy the ActiveSheet to a new workbook
    ActiveSheet.Copy
    Set Destwb = ActiveWorkbook
    
    removeHiddenRows
    'Determine the Excel version and file extension/format
    With Destwb
        If Val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2013
            FileExtStr = ".xlsx": FileFormatNum = 51
        End If
    End With

'    '    'Change all cells in the worksheet to values if you want
'    '    With Destwb.Sheets(1).UsedRange
'    '        .Cells.Copy
'    '        .Cells.PasteSpecial xlPasteValues
'    '        .Cells(1).Select
'    '    End With
'    '    Application.CutCopyMode = False

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = Left(Sourcewb.Name, Len(Sourcewb.Name) - 5) & " " & Format(Now, "dd-mm-yyyy")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    strbody = "<BODY style=font-size:11pt;font-family:Calibri></BODY>"
    
    Application.DisplayAlerts = False
    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = ""
            .Body = ""
            .HTMLBody = strbody
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Display   '.Send or use .Display
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr
     Application.DisplayAlerts = True
    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With
End Sub

Sub removeHiddenRows()
Dim oRow As Range
Dim rng As Range
Dim myRows As Range

With ActiveSheet
    Set myRows = Intersect(.Range("A5:A1048576").EntireRow, .UsedRange)
    If myRows Is Nothing Then Exit Sub
End With

For Each oRow In myRows.Columns(1).Cells
    If oRow.EntireRow.Hidden Then
        If rng Is Nothing Then
            Set rng = oRow
        Else
            Set rng = Union(rng, oRow)
        End If
    End If
Next oRow
Application.DisplayAlerts = False
If Not rng Is Nothing Then rng.EntireRow.delete
Application.DisplayAlerts = True
ActiveCell.AutoFilter
ActiveWorkbook.Sheets(ActiveSheet.Name).Cells.Rows.ClearOutline
End Sub

