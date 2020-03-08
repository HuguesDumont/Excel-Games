VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} perdu 
   Caption         =   "PERDU :("
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   OleObjectBlob   =   "perdu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "perdu"
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

Private Sub recommencer_Click()
    Unload Me
    Unload UserForm1
    UserForm1.Show
End Sub

Private Sub solution_Click()
    Dim i As Long
    
    With UserForm1.Controls
        For i = 1 To 4
            .Item("S" & i).BackColor = .Item(.Item("S" & i).Tag).BackColor
        Next i
    End With
End Sub
