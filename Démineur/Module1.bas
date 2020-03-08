Attribute VB_Name = "Module1"
Option Explicit

Public Sub creerGrille(ligne As Long, col As Long, mines As Long)
    Dim ws As Worksheet, Sh As Worksheet
    Dim i As Long, j As Long, cpt As Long
    Dim plageDem As Range, c As Range, plageVal As Range

    Set ws = ThisWorkbook.Sheets("Démineur")
    Set Sh = ThisWorkbook.Sheets("Valeurs")
    Set plageDem = ws.Range("B2:" & ColLetter(col + 1) & ligne + 1)
    Set plageVal = Sh.Range("B2:" & ColLetter(col + 1) & ligne + 1)
    
    Call bordures.nettoyer
    Call bordures.bord(plageDem)
    Call bordures.bord(plageVal)
    
    For Each c In plageDem.Cells
        c.Interior.Color = RGB(230, 230, 230)
        c.Value = ""
    Next c
    
    For Each c In plageVal.Cells
        c.Interior.Color = RGB(230, 230, 230)
        c.Value = ""
    Next c
    
    Call placerMines(plageVal, mines)
    Call placerNombres(plageVal)
    
    Sh.Cells(1, 64).Value = ligne + 1
    Sh.Cells(1, 65).Value = ColLetter(col + 1)
    Sh.Cells(1, 66).Value = mines
    
    ws.Cells(1, 1).Select
    Sh.Cells(1, 67).Value = Timer
End Sub

Public Function ColLetter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    ColLetter = vArr(0)
End Function

Public Sub placerMines(plage As Range, mines As Long)
    Dim pos As Long
    Dim i As Long
    
    Randomize
    
    For i = 1 To mines
        pos = Int(plage.Cells.Count * Rnd) + 1
        
        While plage.Cells(pos).Value = "X"
            pos = Int(plage.Cells.Count * Rnd) + 1
        Wend
        
        plage.Cells(pos).Value = "X"
    Next i
End Sub

Public Sub placerNombres(plage As Range)
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim c As Range
    
    Set ws = plage.Worksheet
    
    For Each c In plage
        If c.Value <> "X" Then
            ''''''Cellule haut gauche
                If ws.Cells(c.Row - 1, c.Column - 1).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule haut
                If ws.Cells(c.Row - 1, c.Column).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule haut droite
                If ws.Cells(c.Row - 1, c.Column + 1).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule à gauche
                If ws.Cells(c.Row, c.Column - 1).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule à droite
                If ws.Cells(c.Row, c.Column + 1).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule bas gauche
                If ws.Cells(c.Row + 1, c.Column - 1).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule bas
                If ws.Cells(c.Row + 1, c.Column).Value = "X" Then c.Value = c.Value + 1
            ''''''Cellule bas droite
                If ws.Cells(c.Row + 1, c.Column + 1).Value = "X" Then c.Value = c.Value + 1
                
            ''''''Gestion de la couleur
            With c
                If .Value = 1 Then
                    .Font.Color = RGB(0, 192, 192)
                ElseIf .Value = 2 Then
                    .Font.Color = RGB(64, 0, 224)
                ElseIf .Value = 3 Then
                    .Font.Color = RGB(224, 96, 64)
                ElseIf .Value = 4 Then
                    .Font.Color = RGB(255, 64, 0)
                ElseIf .Value = 5 Then
                    .Font.Color = RGB(128, 0, 0)
                ElseIf .Value = 6 Then
                    .Font.Color = RGB(175, 0, 0)
                ElseIf .Value = 7 Then
                    .Font.Color = RGB(210, 0, 0)
                ElseIf .Value = 8 Then
                    .Font.Color = RGB(255, 0, 0)
                Else
                    .Value = 0
                    .Font.Color = RGB(0, 255, 0)
                End If
            End With
        End If
    Next c
End Sub

Sub nouvellePartie()
    param.Show
End Sub
