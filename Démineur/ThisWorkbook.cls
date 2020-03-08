VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''''''Démineur réalisé par Hugues DUMONT
''''''Tous droits réservés
''''''Version 1.0 en date du 16/06/2017

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet
    Dim c As Range, rng As Range
    Dim cpt As Long, ligne As Long
    Dim col As String
    
    If Sh.Name <> "Démineur" Then Exit Sub
    
    If Target.Cells.Count > 1 Then
        Sh.Cells(1, 1).Select
    ElseIf Target.Interior.Color <> RGB(230, 230, 230) Then
        Exit Sub
    Else
        If Target.Value = "" Or Target.Value = "B" Then
            Set ws = ThisWorkbook.Sheets("Valeurs")
            
            Target.Value = ws.Cells(Target.Row, Target.Column).Value
            Target.Font.Color = ws.Cells(Target.Row, Target.Column).Font.Color
            
            ligne = ws.Cells(1, 64).Value
            col = ws.Cells(1, 65).Value
            
            If Target.Value = "X" Then
                Call sons.jouerSon("Bombe")
                Target.Interior.Color = RGB(255, 0, 0)
                perdu.Show
            Else
                cpt = 0
                Set rng = Sh.Range("B2:" & col & ligne)
                For Each c In rng
                    If c.Value = "" Then cpt = cpt + 1
                Next c
                If cpt = ws.Cells(1, 66).Value Then
                    Call sons.jouerSon("Victoire")
                    victoire.Show
                End If
            End If
        End If
    End If
End Sub
