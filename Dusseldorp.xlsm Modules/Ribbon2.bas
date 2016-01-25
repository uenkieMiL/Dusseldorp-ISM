Attribute VB_Name = "Ribbon2"
Option Explicit
Dim ws As String


'Callback for customButton2.1.2 onAction
Sub MaterieelVernieuwen(control As IRibbonControl)
    ws = ActiveSheet.Name
    
    If ws = Blad4.Name Then
            MaterieelPlanningVernieuwen
    End If
End Sub

'Callback for customButton2.1.1 onAction
Sub MaterieelBeheren(control As IRibbonControl)
    ws = ActiveSheet.Name
    
    If ws = Blad4.Name Then
        If IsNumeric(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)) = True Then
            ThisWorkbook.mat_id = ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)
            FORM_MATERIEEL.Show
        End If
    End If
End Sub


'Callback for customButton2.1.3 onAction
Sub MaterieelOverzicht(control As IRibbonControl)
FORM_MATERIEEL_OVERZICHT.Show
End Sub


'Callback for customButton2.2.1 onAction
Sub MaterieelUitgifteGepland(control As IRibbonControl)
    ws = ActiveSheet.Name
    
    If ws = Blad4.Name Then
        If IsNumeric(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)) = True And ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row) <> "" Then
            ThisWorkbook.mat_id = ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)
            FORM_UITGITE_GEPLAND.Show
        End If
    End If
End Sub

'Callback for customButton2.2.2 onAction
Sub MaterieelUitgifteOverig(control As IRibbonControl)
FORM_UITGITE_OVERIG.Show
End Sub

'Callback for customButton2.2.3 onAction
Sub MaterieelInname(control As IRibbonControl)
Dim m As New Materieel
ws = ActiveSheet.Name
Dim ma As New MaterieelActie
If ThisWorkbook.Sheets(ws).Range(MaterielenPlanning.col_mat_status & ActiveCell.Row) = "In Magazijn" Then Exit Sub
    
    If ws = Blad4.Name Then
        If IsNumeric(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row)) = True Then
            m.Id = CLng(ThisWorkbook.Sheets(ws).Range("A" & ActiveCell.Row).Value)
            m.haalop
            m.Status = "In Magazijn"
            If m.updateStatus = True Then
                ThisWorkbook.Sheets(ws).Range(MaterielenPlanning.col_mat_status & ActiveCell.Row).Value = "In Magazijn"
                ma.MaterieelId = m.Id
                ma.Waarde = "In Magazijn"
                ma.InsertInname
            Else
                Functies.errorhandler_MsgBox ("Er is iets misgegaan met aanpassen van de status.")
            End If
        End If
    End If
End Sub

'Callback for cu    stomButton2.2.4 onAction
Sub MaterieelInnameOverig(control As IRibbonControl)
    FORM_INNAME_OVERIG.Show
End Sub

'Callback for customButton2.3.1 onAction
Sub MaterieelNieweBestellingen(control As IRibbonControl)
FORM_BESTELLINGEN.Show
End Sub

'Callback for customButton2.3.2 onAction
Sub OrdersPicken(control As IRibbonControl)
End Sub

