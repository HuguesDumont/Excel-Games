VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Memory 
   Caption         =   "Memory familial"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   OleObjectBlob   =   "Memory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub affExcel_Click()
    Unload Me
    Application.Visible = True
End Sub

Private Sub NewGame_Click()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Activate()
    Call lesMenus.AddMinimiseButton
End Sub

Private Sub UserForm_Initialize()
    Dim ctrl As Control
    Dim i As Integer, carte As Integer, cpt As Integer
    Dim cartes(9) As Integer
    Dim fin As Boolean
    
    For Each ctrl In Me.Controls
        ctrl.Caption = ""
        ctrl.BackColor = "&HFF8080"
        ctrl.Picture = Nothing
    Next ctrl
    
    Me.Caption = "Memory familial"
    Me.BackColor = RGB(255, 255, 255)
    
    Me.NewGame.Caption = "Nouvelle partie"
    Me.NewGame.BackColor = RGB(0, 50, 255)
    Me.affExcel.Caption = "Afficher Excel"
    Me.affExcel.BackColor = RGB(0, 200, 0)
    Me.Label1.BackColor = RGB(255, 255, 255)
    Randomize
    
    ''''''''Initialisation du plateau à zéro pour éviter les rendus avec des valeurs inattendues de la mémoire
    For i = 0 To 19
        Module1.plateau(i) = -1
    Next i
    
    fin = False
    cpt = 0
    
    For i = 0 To 9
        cartes(i) = 0
        Module1.ima(i) = ThisWorkbook.Path & "\" & i & ".jpg"
    Next i
    
    ''''''On attribue les cartes aléatoirement
    While Not fin
        i = Int(20 * Rnd)
        If plateau(i) < 0 Then
            carte = Int(10 * Rnd)
            If cartes(carte) < 2 Then
                plateau(i) = carte
                cartes(carte) = cartes(carte) + 1
                cpt = cpt + 1
            End If
        End If
        If cpt = 20 Then fin = True
    Wend

    For i = 0 To 19
        Me.Controls.Item("Image" & i + 1).Enabled = True
    Next i
    
    Module1.previous = -1
    Module1.trouvees = 0
    Module1.nbCoups = 0
    Me.Label1.Caption = "Coups joués : " & Module1.nbCoups
End Sub

Private Sub Image1_Click()
    Call jouer(0)
End Sub

Private Sub Image2_Click()
    Call jouer(1)
End Sub

Private Sub Image3_Click()
    Call jouer(2)
End Sub

Private Sub Image4_Click()
    Call jouer(3)
End Sub

Private Sub Image5_Click()
    Call jouer(4)
End Sub

Private Sub Image6_Click()
    Call jouer(5)
End Sub

Private Sub Image7_Click()
    Call jouer(6)
End Sub

Private Sub Image8_Click()
    Call jouer(7)
End Sub

Private Sub Image9_Click()
    Call jouer(8)
End Sub

Private Sub Image10_Click()
    Call jouer(9)
End Sub

Private Sub Image11_Click()
    Call jouer(10)
End Sub

Private Sub Image12_Click()
    Call jouer(11)
End Sub

Private Sub Image13_Click()
    Call jouer(12)
End Sub

Private Sub Image14_Click()
    Call jouer(13)
End Sub

Private Sub Image15_Click()
    Call jouer(14)
End Sub

Private Sub Image16_Click()
    Call jouer(15)
End Sub

Private Sub Image17_Click()
    Call jouer(16)
End Sub

Private Sub Image18_Click()
    Call jouer(17)
End Sub

Private Sub Image19_Click()
    Call jouer(18)
End Sub

Private Sub Image20_Click()
    Call jouer(19)
End Sub
