Attribute VB_Name = "GlobalVars"
Option Explicit

Public li_NbCoups As Integer ' Nombre de coups effectuées
Public li_CoupsBons As Integer ' Nombre de taupes touchées
Public ld_Speed As Double ' Vitesse du jeu (durée d'affichage max d'une "taupe")
Public ls_Col As String ' ls_Colonne de correspondance à la vitesse
Public ld_Duree As Double ' Durée de la partie
Public ls_PlayerName As String ' Variable contenant le nom du joueur
Public ll_PlayerLine As Long ' Variable contenant la ligne où se trouve le joueur (pour la mise à jour des résultats)
