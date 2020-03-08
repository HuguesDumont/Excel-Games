VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} game 
   Caption         =   "Jeu du chasse-taupe par Hugues DUMONT"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   OleObjectBlob   =   "game.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Cases pouvant contenir des taupes
Private Sub CommandButton1_Click()
    Call verify(1)
End Sub

Private Sub CommandButton2_Click()
    Call verify(2)
End Sub

Private Sub CommandButton3_Click()
    Call verify(3)
End Sub

Private Sub CommandButton4_Click()
    Call verify(4)
End Sub

Private Sub CommandButton5_Click()
    Call verify(5)
End Sub

Private Sub CommandButton6_Click()
    Call verify(6)
End Sub

Private Sub CommandButton7_Click()
    Call verify(7)
End Sub

Private Sub CommandButton8_Click()
    Call verify(8)
End Sub

Private Sub CommandButton9_Click()
    Call verify(9)
End Sub

Private Sub CommandButton10_Click()
    Call verify(10)
End Sub

Private Sub CommandButton11_Click()
    Call verify(11)
End Sub

Private Sub CommandButton12_Click()
    Call verify(12)
End Sub

Private Sub CommandButton13_Click()
    Call verify(13)
End Sub

Private Sub CommandButton14_Click()
    Call verify(14)
End Sub

Private Sub CommandButton15_Click()
    Call verify(15)
End Sub

Private Sub CommandButton16_Click()
    Call verify(16)
End Sub

Private Sub CommandButton17_Click()
    Call verify(17)
End Sub

Private Sub CommandButton18_Click()
    Call verify(18)
End Sub

Private Sub CommandButton19_Click()
    Call verify(19)
End Sub

Private Sub CommandButton20_Click()
    Call verify(20)
End Sub

Private Sub CommandButton21_Click()
    Call verify(21)
End Sub

Private Sub CommandButton22_Click()
    Call verify(22)
End Sub

Private Sub CommandButton23_Click()
    Call verify(23)
End Sub

Private Sub CommandButton24_Click()
    Call verify(24)
End Sub

Private Sub CommandButton25_Click()
    Call verify(25)
End Sub

' Commencer une nouvelle partie
Private Sub newGame_Click()
    Dim debut As Double, fin As Double ' Durée de la partie
    Dim taupeDebut As Double, taupeFin As Double ' Durée de l'affichage d'une "taupe"
    Dim choixTaupe As Integer ' Choix de la case contenant la "taupe"
    Dim i As Long ' Variable de boucle (pour le parcours des item du usf)

    Randomize
    
    Me.NewGame.Enabled = False ' Bloquer le lancement d'une partie alors qu'une partie est déjà en cours
    Me.TempsRestant.Caption = CStr(ld_Duree) ' Affichage de la durée de la partie
    
    ' Intialiser le jeu
    For i = 1 To 25
        With Me.Controls.Item("CommandButton" & i)
            .Enabled = True ' débloquer les boutons pour taper sur les "taupes"
            .Tag = "False" ' Intiliaser les "taupes" à Faux
        End With
    Next i
        
    debut = Timer ' début de la partie
    fin = debut + 0.0000001 ' Borne pour vérifier le temps restant à jouer
    
    While ((fin - debut) < GlobalVars.ld_Duree) ' tant que le temps de jeu n'a pas dépassé la durée max de jeu
        choixTaupe = Int(Rnd * 25) + 1 ' On choisi une case pour la "taupe"
        
        With Me.Controls.Item("CommandButton" & choixTaupe)
            .Tag = "True" ' A laquelle on indique qu'il y a une taupe
            .BackColor = RGB(255, 0, 0) ' Et on affiche la taupe
        End With
        
        taupeDebut = Timer ' On récupère le moment du début d'affichage de la taupe
        taupeFin = Timer ' Borne pour vérifier le temps restant d'affichage de la taupe
        
        ' Tant que la taupe est affichée et qu'elle n'a pas été tapée (fonction verify)
        While (((taupeFin - taupeDebut < GlobalVars.ld_Speed) And (Me.Controls.Item("CommandButton" & choixTaupe).BackColor = RGB(255, 0, 0))) And (fin - debut < GlobalVars.ld_Duree))
            Me.decompte.Width = (fin - debut) * 252 / 60 ' Mettre à jour l'affichage de la "Barre de progression"
            Me.TempsRestant.Caption = 60 - Round(fin - debut, 0) ' Affichage du temps restant
            taupeFin = Timer ' Borne pour vérifier le temps restant d'affichage de la taupe
            fin = Timer ' Borne pour vérifier le temps restant à jouer
            DoEvents ' Autoriser les autres actions (Click et utilisations d'autres fonctions)
        Wend
        
        ' Le temps d'affichage de la taupe est dépassé donc
        With Me.Controls.Item("CommandButton" & choixTaupe)
            .Tag = "False" ' On indique qu'il n'y a plus de taupe
            .BackColor = -2147483633 ' On cache la taupe
        End With
        DoEvents ' Autoriser les autres actions
    Wend

    Unload Me ' Fermer la fenêtre de jeu
    endGame.Show ' Afficher la fenêtre de résultat
End Sub

' Vérifier la présence ou non d'une taupe à l'endroit du click
Private Sub verify(ByVal taupe As Integer)
    With Me.Controls.Item("CommandButton" & taupe)
        If CBool(.Tag) Then ' Si une taupe est présente
            GlobalVars.li_CoupsBons = GlobalVars.li_CoupsBons + 1 ' On augmente le nombre de coups corrects
            .Tag = "False" ' On indique qu'il n'y a plus de taupe
            .BackColor = -2147483633 ' On retire la taupe
        End If
    End With
    
    GlobalVars.li_NbCoups = GlobalVars.li_NbCoups + 1 ' Augmenter le nombre de coups donnés à chaque click (correct ou incorrect)
    DoEvents ' Permettre les autres actions
End Sub

Private Sub UserForm_Initialize()
    Me.TempsRestant.Caption = CStr(GlobalVars.ld_Duree)
    Me.Label1.Caption = "Frappez le plus de taupes possibles en " & GlobalVars.ld_Duree & " secondes"
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
