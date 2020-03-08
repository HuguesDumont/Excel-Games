VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   12630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim plateau(9, 9) As Frame
    Dim i As Integer, j As Integer
    Dim lettres(9) As String, nom As String
    
    For i = 0 To 9
        lettres(i) = Chr(i + 65)
    Next i
    
    For i = 0 To 9
        'Initialisation des blancs
        For j = 0 To 3
            If (j Mod 2 = 0) Then
                If (i Mod 2 <> 0) Then
                    nom = lettres(i) & (j + 1)
                    Me.Controls.Item(nom).Picture = LoadPicture("\\f-mout\home2$\a000678\MyDocs\Travail Renault\Utilitaires programmation\Playing\A Finir\Dames\Blanc.gif")
                End If
            Else
                If (i Mod 2 = 0) Then
                    nom = lettres(i) & (j + 1)
                    Me.Controls.Item(nom).Picture = LoadPicture("\\f-mout\home2$\a000678\MyDocs\Travail Renault\Utilitaires programmation\Playing\A Finir\Dames\Blanc.gif")
                End If
            End If
        Next j
        'Initialisation des noirs
        For j = 6 To 9
            If (j Mod 2 = 0) Then
                If (i Mod 2 <> 0) Then
                    nom = lettres(i) & (j + 1)
                    Me.Controls.Item(nom).Picture = LoadPicture("\\f-mout\home2$\a000678\MyDocs\Travail Renault\Utilitaires programmation\Playing\A Finir\Dames\Noir.gif")
                End If
            Else
                If (i Mod 2 = 0) Then
                    nom = lettres(i) & (j + 1)
                    Me.Controls.Item(nom).Picture = LoadPicture("\\f-mout\home2$\a000678\MyDocs\Travail Renault\Utilitaires programmation\Playing\A Finir\Dames\Noir.gif")
                End If
            End If
        Next j
    Next i
End Sub

Private Sub B1_Click()

End Sub

Private Sub D1_Click()

End Sub

Private Sub J1_Click()

End Sub

Private Sub F1_Click()

End Sub

Private Sub H1_Click()

End Sub

Private Sub A2_Click()

End Sub

Private Sub C2_Click()

End Sub

Private Sub E2_Click()

End Sub

Private Sub G2_Click()

End Sub

Private Sub I2_Click()

End Sub

Private Sub A4_Click()

End Sub

Private Sub B3_Click()

End Sub

Private Sub D3_Click()

End Sub

Private Sub J3_Click()

End Sub

Private Sub F3_Click()

End Sub

Private Sub H3_Click()

End Sub

Private Sub C4_Click()

End Sub

Private Sub E4_Click()

End Sub

Private Sub G4_Click()

End Sub

Private Sub I4_Click()

End Sub

Private Sub B5_Click()

End Sub

Private Sub D5_Click()

End Sub

Private Sub J5_Click()

End Sub

Private Sub F5_Click()

End Sub

Private Sub H5_Click()

End Sub

Private Sub A6_Click()

End Sub

Private Sub C6_Click()

End Sub

Private Sub E6_Click()

End Sub

Private Sub G6_Click()

End Sub

Private Sub I6_Click()

End Sub

Private Sub A8_Click()

End Sub

Private Sub B7_Click()

End Sub

Private Sub D7_Click()

End Sub

Private Sub J7_Click()

End Sub

Private Sub F7_Click()

End Sub

Private Sub H7_Click()

End Sub

Private Sub C8_Click()

End Sub

Private Sub E8_Click()

End Sub

Private Sub G8_Click()

End Sub

Private Sub I8_Click()

End Sub

Private Sub A10_Click()

End Sub

Private Sub B9_Click()

End Sub

Private Sub D9_Click()

End Sub

Private Sub J9_Click()

End Sub

Private Sub F9_Click()

End Sub

Private Sub H9_Click()

End Sub

Private Sub C10_Click()

End Sub

Private Sub E10_Click()

End Sub

Private Sub G10_Click()

End Sub

Private Sub I10_Click()

End Sub
