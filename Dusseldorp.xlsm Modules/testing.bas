Attribute VB_Name = "testing"
Function testnewkalendar()

SoortPlanning.MaakSoortPlanning Blad2.Name, 2
End Function

Sub testformwijzigen()
ThisWorkbook.synergy_id = ThisWorkbook.Sheets(Blad3.Name).Range("A" & ActiveCell.Row)
ThisWorkbook.Vestiging = ThisWorkbook.Sheets(Blad3.Name).Range("B" & ActiveCell.Row)
FORM_PROJECT_WIJZIGEN.Show
End Sub

Sub testMaterieel()
Dim mp As New MaterieelPlanning
mp.Id = 20
mp.GetById
mp.Print_r

End Sub


