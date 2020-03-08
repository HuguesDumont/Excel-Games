VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} optionSelector 
   Caption         =   "Choix des options"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   OleObjectBlob   =   "optionSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "optionSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Commencer_Click()
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    Dim trouve As Boolean
    
    Set ws = ThisWorkbook.Sheets("Individuels")
    
    i = 2
    trouve = False
    
    While ((i < ws.Cells(Rows.Count, 1).End(xlUp).Row) And (Not trouve)) ' Boucle de recherche du pseudo
        If (ws.Cells(i, 1).Value = Me.player.Value) Then ' Cas où le pseudo est trouvé
            For j = 2 To 9 ' Boucle d'affichage des scores
                Me.Controls.Item("v" & j - 1).Caption = ws.Cells(i, j).Value ' Affichage des scores pour la vitesse correspondante
            Next j
             ' Pseudo trouvé donc on quitte la boucle (tous les pseudos sont ainsi uniques)
            trouve = True
        End If
        i = i + 1
        DoEvents
    Wend
    
    If (i > ws.Cells(Rows.Count, 1).End(xlUp).Row) Then
        ' Ajouter le joueur à la liste des joueurs
        ws.Cells(i, 1).Value = Me.player.Value
    End If
    
    ' Récupérer le nom et la ligne correspondant à celle du joueur (pour la mise à jour des résultats par la suite)
    GlobalVars.ls_PlayerName = Me.player.Value
    GlobalVars.ll_PlayerLine = i
    
    Unload Me
    game.Show
End Sub

' Affichage des meilleurs scores si le pseudo du joueur est trouvé
Private Sub player_Change()
    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim trouve As Boolean
    
    Set ws = ThisWorkbook.Sheets("Individuels")
    
    i = 2
    trouve = False
    
    While ((i < ws.Cells(Rows.Count, 1).End(xlUp).Row) And (Not trouve)) ' Boucle de recherche du pseudo
        If (ws.Cells(i, 1).Value = Me.player.Value) Then ' Cas où le pseudo est trouvé
            For j = 2 To 9 ' Boucle d'affichage des scores
                Me.Controls.Item("v" & j - 1).Caption = ws.Cells(i, j).Value ' Affichage des scores pour la vitesse correspondante
            Next j
             ' Pseudo trouvé donc on quitte la boucle (tous les pseudos sont ainsi uniques)
            trouve = True
        End If
        i = i + 1
        DoEvents
    Wend
    
    ' Mettre à jour l'affichage des résultats individuels
    For j = 2 To 9
        Me.Controls.Item("v" & j - 1).Caption = ws.Cells(i, j).Value
    Next j
End Sub

Private Sub tresLente_Click()
    Call updating(3, "B")
End Sub

Private Sub Lente_Click()
    Call updating(2, "D")
End Sub

Private Sub Ralentie_Click()
    Call updating(1.5, "F")
End Sub

Private Sub Normale_Click()
    Call updating(1, "H")
End Sub

Private Sub Acceleree_Click()
    Call updating(0.75, "J")
End Sub

Private Sub Rapide_Click()
    Call updating(0.5, "L")
End Sub

Private Sub TresRapide_Click()
    Call updating(0.25, "N")
End Sub

Private Sub Impossible_Click()
    Call updating(0.1, "P")
End Sub

' Fonction de mise à jour des meilleurs scores pour la vitesse sélectionnée
Private Sub updating(ByVal vitesse As Double, ByVal ls_Colonne As String)
    Dim i As Long
    Dim ws As Worksheet
    Dim rng As Range
    
    GlobalVars.ls_Col = ls_Colonne ' Indiquer la colonne correspondant à la vitesse sélectionnée
    
    ' Trier les meilleurs scores dans l'ordre décroissant pour la vitesse concernée
    Set ws = ThisWorkbook.Sheets("Records")
    Set rng = ws.Range(Chr(Asc(ls_Colonne) - 1) & "1:" & ls_Colonne & "11")
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ls_Colonne & "2:" & ls_Colonne & "11"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    GlobalVars.ld_Speed = vitesse ' Indication de la vitesse du jeu
    
    ' Affichage des meilleurs scores
    Me.scores.Clear
    
    For i = 2 To 11
        Me.scores.AddItem
        Me.scores.List(Me.scores.ListCount - 1, 0) = ws.Cells(i, Asc(ls_Colonne) - 65).Value
        Me.scores.List(Me.scores.ListCount - 1, 1) = ws.Cells(i, Asc(ls_Colonne) - 64).Value
    Next i
End Sub

Private Sub UserForm_Initialize()
    ' Initialisation des valeurs à la vitesse normale
    Me.Normale.Value = True
    
    ' Récupération du pseudo du précédent joueur (si le jeu n'a pas été fermé)
    Me.player.Value = GlobalVars.ls_PlayerName
    
    ' Pour mises à jours futures (sélection de la durée de la partie
    GlobalVars.ld_Duree = 60
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim li_cpt As Integer
    
    If (CloseMode = 1) Then
        Cancel = True
    ElseIf (Application.Workbooks.Count > 1) Then
        Unload Me
        ThisWorkbook.Close SaveChanges:=False
    Else
        Application.DisplayAlerts = False
        
        For li_cpt = 1 To Application.AddIns2.Count
            Application.AddIns2(li_cpt).Application.Quit
        Next li_cpt
        
        For li_cpt = 1 To Application.AddIns.Count
            Application.AddIns(li_cpt).Application.Quit
        Next li_cpt
        
        For li_cpt = 1 To Application.COMAddIns.Count
            Application.COMAddIns(li_cpt).Application.Quit
        Next li_cpt
        
        Unload Me
        ThisWorkbook.Close False
        
        Application.Quit
    End If
End Sub
