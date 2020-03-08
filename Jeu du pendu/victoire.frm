VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} victoire 
   Caption         =   "Victoire !   :D"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   OleObjectBlob   =   "victoire.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "victoire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fermer_Click()
    Unload Me
End Sub

Private Sub quitter_Click()
    Unload Me
    ThisWorkbook.Close
    Application.Quit
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim cpt As Long
    
    cpt = 0
    With pendu.tried
        For i = 1 To Len(.Caption)
            If Mid(.Caption, i, 1) <> "-" And Mid(.Caption, i, 1) <> " " And Mid(.Caption, i, 1) <> "_" Then
                cpt = cpt + 1
            End If
        Next i
    End With
    Me.Label1.Caption = "BRAVO ! Vous avez gagné avec " & cpt & " lettres."
End Sub
