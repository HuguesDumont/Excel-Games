Attribute VB_Name = "Module1"
Option Explicit

Public premier As Boolean
Public ligne As Long, colonne As Long

Public Sub game()

    Randomize
    
    ligne = Int(15 * Rnd) + 1
    colonne = Int(39 * Rnd) + 1
    
    premier = True
    Cells.Select
    Selection.ColumnWidth = 4
    Selection.RowHeight = 24
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Public Sub verif()
    Dim l As Integer, c As Integer
    l = Selection.Row
    c = Selection.Column
    If (l = ligne And c = colonne) Then
        MsgBox ("gagné")
        Selection.Interior.Color = RGB(0, 255, 0)
        Call game
    ElseIf ((Abs(l - ligne) <= 1) And (Abs(c - colonne) <= 1)) Then
        MsgBox ("Chaud bouillant")
        Selection.Interior.Color = RGB(150, 210, 150)
    ElseIf ((Abs(l - ligne) <= 3) And (Abs(c - colonne) <= 3)) Then
        MsgBox ("Très chaud")
        Selection.Interior.Color = RGB(0, 175, 100)
    ElseIf ((Abs(l - ligne) <= 5) And (Abs(c - colonne) <= 5)) Then
        MsgBox ("Ca se réchauffe")
        Selection.Interior.Color = RGB(100, 175, 0)
    ElseIf ((Abs(l - ligne) <= 8) And (Abs(c - colonne) <= 8)) Then
        MsgBox ("C'est froid")
        Selection.Interior.Color = RGB(100, 120, 0)
    Else
        MsgBox ("AGLAGLAGLAGLAGLAGLA")
        Selection.Interior.Color = RGB(0, 0, 255)
    End If
End Sub
