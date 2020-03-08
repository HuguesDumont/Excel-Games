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

Private Sub quitter_Click()
    Unload Me
    ThisWorkbook.Close
    Application.Quit
End Sub

Private Sub newGame_Click()
    Unload Me
    param.Show
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Valeurs")
    Me.Label1.Caption = "BRAVO ! Vous avez gagné en " & Format((Timer - ws.Cells(1, 67).Value) / 86400, "hh:mm:ss")
End Sub
