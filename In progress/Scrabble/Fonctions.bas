Attribute VB_Name = "Fonctions"
Option Explicit

Public Function sumCol(f As Worksheet, c As Integer, Optional debut As Long = 1, Optional fin As Long = -1) As Long
    Dim i As Long
    
    sumCol = 0
    If fin = -1 Then fin = f.Cells(debut, c).End(xlDown).Row
    
    For i = debut To fin
        sumCol = sumCol + f.Cells(i, c).Value
    Next i
End Function

Public Sub initGame(col As Long, ligneAjout As Long)
    Dim i As Long, j As Integer
    
    With Variables.fGame
        If (.Cells(3, 1).Value <> "") Then
            For i = 3 To .Cells(3, 1).End(xlDown).Row
                .Cells(i, 1).Value = ""
                .Cells(i, 2).Value = ""
            Next i
        End If
        For i = 1 To 4
            .Cells(2 * i + 1, 4).Value = 0
            For j = 1 To 7
                .Cells(2 * i + 1, 4 + j).Value = ""
                .Cells(2 * i + 2, 4 + j).Value = ""
            Next j
        Next i
    End With
    
    With Variables.fPions
        For i = 4 To .Cells(4, col).End(xlDown).Row
            j = .Cells(i, col + 2).Value
            While j <> 0
                Variables.fGame.Cells(ligneAjout, 1).Value = .Cells(i, col).Value
                Variables.fGame.Cells(ligneAjout, 2).Value = .Cells(i, col + 1).Value
                ligneAjout = ligneAjout + 1
                j = j - 1
            Wend
        Next i
    End With
End Sub

Public Sub distrib(joueur As Integer)
    Dim i As Integer, j As Integer
    Dim sup As Long
    
    Randomize
    
    With Variables.fGame
        For i = 1 To 7
            sup = .Cells(3, 1).End(xlDown).Row
            If .Cells(3, 1).Value <> "" Then
                j = Int(sup * Rnd) + 3
                .Cells(joueur, 4 + i).Value = .Cells(j, 1).Value
                .Cells(joueur + 1, 4 + i).Value = .Cells(j, 2).Value
                .Range(.Cells(j, 1), .Cells(j, 2)).Delete Shift:=xlUp
            End If
        Next i
    End With
End Sub

Public Function verif(mot As String) As Boolean
    Dim c As Range, col As Range, apres As Range
    
    Set col = Variables.fDic.Columns(Variables.numLangue)
    Set apres = Variables.fDic.Cells(MMMMMMM, Variables.numLangue)
    Set c = col.Find(mot, apres, , xlWhole, xlByRows, xlNext, False, False)
    
    verif = Not (c Is Nothing)
End Function

''''''Fonction de déplacements spéciaux pour listbox et combobox, On retourne la valeur que prendra keycode pour empêcher les déplacements en surplus
Public Function scrollFleche(usf As UserForm, leNom As String, ByVal laFleche As Integer) As Integer
    With usf.Controls.Item(leNom)
        scrollFleche = laFleche
        If laFleche = 40 Then ''''''Flèche bas
            If .ListIndex = .ListCount - 1 Then ''''''Si élément sélectionné est le dernier alors
                .ListIndex = 0  ''''''sélectionner le premier
                scrollFleche = 0 ''''''Et empêcher le double-déplacement
            End If
        ElseIf laFleche = 38 Then ''''''Flèche haut
            If .ListIndex = 0 Then ''''''Si élément sélectionné est le premier alors
                .ListIndex = .ListCount - 1 ''''''sélectionner le dernier
                scrollFleche = 0 ''''''Et empêcher le double-déplacement
            End If
        End If
    End With
End Function
''''''
