Attribute VB_Name = "WeekPlanning"
Public Const skw = 11

Function MaakWekenlijst(wacht As Boolean, Vestiging As String)
Dim weekvandaag As Long: weekvandaag = -1
Application.DisplayAlerts = False
Dim weeklijst As Variant: weeklijst = WeekOverzichtOphalen(wacht, Vestiging)
Dim lk As Integer: lk = ThisWorkbook.Sheets(Blad7.Name).Range("XFD3").End(xlToLeft).Column
ThisWorkbook.Sheets(Blad7.Name).Range(Functies.KolomNaarLetter(skw) & ":ZZ").Clear
ThisWorkbook.Sheets(Blad7.Name).Range(Functies.KolomNaarLetter(skw) & ":ZZ").ColumnWidth = 3


ThisWorkbook.Sheets(Blad7.Name).Range("A4:J1048576").Clear
If (weeklijst(1, UBound(weeklijst, 2)) - weeklijst(1, UBound(weeklijst, 2) - 1)) <= 1 Then laatsteweek = 0 Else laatsteweek = 1
For k = 0 To UBound(weeklijst, 2) - laatsteweek
ThisWorkbook.Sheets(Blad7.Name).Range(Cells(1, skw + k).Address) = weeklijst(0, k)
ThisWorkbook.Sheets(Blad7.Name).Range(Cells(2, skw + k).Address) = Maandnaam(CLng(weeklijst(2, k)))
ThisWorkbook.Sheets(Blad7.Name).Range(Cells(3, skw + k).Address) = weeklijst(1, k)
If weekvandaag = -1 And weeklijst(1, k) = IsoWeekNumber(Now()) Then weekvandaag = k
Next k
Blad7.Visible = xlSheetVisible
Blad7.Select

If weekvandaag > -1 Then ThisWorkbook.Sheets(Blad7.Name).Range(Cells(4, skw + weekvandaag), Cells(1048576, skw + weekvandaag)).Interior.Color = 12611584

Call WeekPlanning.WeekJaarSamenvoegen
Call WeekPlanning.WeekMaandSamenvoegen

    With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(3, 1), Cells(3, lk)).Address).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
Application.DisplayAlerts = True
End Function
Function WeekJaarSamenvoegen()
Dim lk As Long
Dim k As Long
Dim startrij As Long: startrij = skw

lk = ThisWorkbook.Sheets(Blad7.Name).Range("XFD1").End(xlToLeft).Column
Application.DisplayAlerts = False
For k = skw + 1 To lk + 1

    If ThisWorkbook.Sheets(Blad7.Name).Range(Cells(1, k).Address) <> ThisWorkbook.Sheets(Blad7.Name).Range(Cells(1, k - 1).Address) Then
    
    With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(1, startrij), Cells(1, k - 1)).Address)
        .Merge
        .HorizontalAlignment = xlCenter
    End With
    
        With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(1, startrij), Cells(1, k - 1)).Address).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(1, startrij), Cells(1, k - 1)).Address).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(1, startrij), Cells(1, k - 1)).Address).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(1, startrij), Cells(1, k - 1)).Address).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    startrij = k
End If

Next k
Application.DisplayAlerts = True
End Function


Function WeekMaandSamenvoegen()
Dim lk As Long
Dim sk As Long
Dim k As Long
Dim startrij As Long: startrij = skw
lk = ThisWorkbook.Sheets(Blad7.Name).Range("XFD2").End(xlToLeft).Column
sk = skw
Application.DisplayAlerts = False
For k = skw + 1 To lk + 1
    If ThisWorkbook.Sheets(Blad7.Name).Range(Cells(2, k).Address) <> ThisWorkbook.Sheets(Blad7.Name).Range(Cells(2, k - 1).Address) Then
        
    
        With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Columns(startrij), Columns(k - 1)).Address).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    
        With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Columns(startrij), Columns(k - 1)).Address).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        With ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(2, startrij), Cells(2, k - 1)).Address)
            .Merge
            .HorizontalAlignment = xlCenter
        End With
'
'        With ThisWorkbook.Sheets(blad7.Name).Range(Range(Cells(, sk), Cells(0, k - 1)).Address)
'        .LineStyle = xlContinuous
'        .ColorIndex = xlAutomatic
'        .TintAndShade = 0
'        .Weight = xlMedium
'        End With

    
        startrij = k
    End If

Next k
Application.DisplayAlerts = True
End Function

Function DatumNaarKolomWeek(datum As Date, lijst As Variant) As Long


    For k = 0 To UBound(lijst, 2)
        If lijst(0, k) = Year(datum) And IsoWeekNumber(datum) = lijst(1, k) Then
            DatumNaarKolomWeek = skw + k
            Exit Function
        Else
        DatumNaarKolomWeek = -1
        End If
    Next k


End Function

Function RijenKolommenNaarRange(rij As Long, k1 As Long, k2 As Long) As Range
If k1 = 0 Then k1 = skw
If k2 = 0 Then k2 = skw
If k1 > -1 And k2 > -1 Then
    Set RijenKolommenNaarRange = ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(rij, k1), Cells(rij, k2)).Address)
ElseIf k1 = -1 And k2 = -1 Then
    Set RijenKolommenNaarRange = Nothing
Else
If k1 = -1 Then k1 = skw
If k2 = -1 Then k2 = skw

 Set RijenKolommenNaarRange = ThisWorkbook.Sheets(Blad7.Name).Range(Range(Cells(rij, k1), Cells(rij, k2)).Address)
End If
End Function

Function WeekOverzichtOphalen(wacht As Boolean, Vestiging As String) As Variant
Dim sql As String
Dim subquery As String


If Vestiging <> "" And wacht = True And Status = True Then
    Vestiging = " WHERE Vestiging = '" & Vestiging & "'"
ElseIf Vestiging <> "" Then
    Vestiging = " AND Vestiging = '" & Vestiging & "'"
End If



If wacht = True Then
    subquery = "select Synergy from PROJECTEN WHERE STATUS = 0" & Vestiging
Else
    subquery = "select Synergy from PROJECTEN WHERE STATUS = 0 AND WACHT = 0" & Vestiging
End If

sql = "SELECT DISTINCT " & _
"DatePart('yyyy',[DAGEN].[DATUMNIEUW],2,2) AS JAAR, " & _
"DatePart('ww',[DAGEN].[DATUMNIEUW],2,2) AS WEEK, " & _
"Min(DatePart('m', [dagen].[DATUMNIEUW],2,2)) As MAAND " & _
"FROM (SELECT A.DATUM AS " & _
        "DATUMNIEUW, " & _
        "A.FEESTDAG, " & _
        "A.EXTRADAG, " & _
        "A.OMSCHRIJVING, " & _
        "A.ZICHTBAAR, " & _
        "(SELECT COUNT(*) FROM KALENDER WHERE A.DATUM>=KALENDER.DATUM) AS KOLOMNUMMER " & _
        "FROM KALENDER AS A " & _
        "WHERE A.DATUM>=(select min(Startdatum) from PRODUCTIE WHERE Synergy IN (" & subquery & ")) " & _
        " And A.DATUM<=(select max(einddatum) from PRODUCTIE WHERE Synergy IN (" & subquery & ")) " & _
"ORDER BY A.DATUM) DAGEN " & _
"GROUP BY DatePart('yyyy',[DAGEN].[DATUMNIEUW],2,2), DatePart('ww',[DAGEN].[DATUMNIEUW],2,2) " & _
"ORDER BY DatePart('yyyy',[DAGEN].[DATUMNIEUW],2,2), DatePart('ww',[DAGEN].[DATUMNIEUW],2,2);"


WeekOverzichtOphalen = LijstOpBasisVanQuery(sql)


End Function


