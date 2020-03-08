VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} param 
   Caption         =   "Démineur (Hugues) - Paramètres du jeu"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   OleObjectBlob   =   "param.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MineInf As Long, MineSup As Long

Private Sub nbLignes_AfterUpdate()
    With Me.nbLignes
        If Not IsNumeric(.Value) Or .Value = "" Then
            MsgBox "La valeur saisie n'est pas un nombre.", vbOKOnly + vbExclamation, "Valeur incorrecte"
            .Value = 10
        ElseIf .Value < 5 Or .Value > 20 Then
            MsgBox "Nombre de lignes incorrect.", vbOKOnly + vbExclamation, "Nombre de lignes incorrect"
            .Value = 10
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub nbCol_AfterUpdate()
    With Me.nbCol
        If Not IsNumeric(.Value) Or .Value = "" Then
            MsgBox "La valeur saisie n'est pas un nombre.", vbOKOnly + vbExclamation, "Valeur incorrecte"
            .Value = 10
        ElseIf .Value < 5 Or .Value > 50 Then
            MsgBox "Nombre de colonnes incorrect.", vbOKOnly + vbExclamation, "Nombre de colonnes incorrect"
            .Value = 10
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub SpinLigne_SpinDown()
    With Me.nbLignes
        If .Value = 5 Then
            MsgBox "Nombre de lignes minimum atteint.", vbOKOnly + vbExclamation, "Minimum atteint"
        Else
            .Value = .Value - 1
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub SpinLigne_SpinUp()
    With Me.nbLignes
        If .Value = 20 Then
            MsgBox "Nombre de lignes maximum atteint.", vbOKOnly + vbExclamation, "Maximum atteint"
        Else
            .Value = .Value + 1
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub SpinCol_SpinDown()
    With Me.nbCol
        If .Value = 5 Then
            MsgBox "Nombre de lignes minimum atteint.", vbOKOnly + vbExclamation, "Minimum atteint"
        Else
            .Value = .Value - 1
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub Spincol_SpinUp()
    With Me.nbCol
        If .Value = 50 Then
            MsgBox "Nombre de lignes maximum atteint.", vbOKOnly + vbExclamation, "Maximum atteint"
        Else
            .Value = .Value + 1
        End If
    End With
    Me.nbCases.Caption = CStr(Me.nbLignes.Value * Me.nbCol.Value)
    Me.nbMines.Value = CInt(CInt(Me.nbCases.Caption) / 5)
    MineInf = CInt(CInt(Me.nbCases.Caption) / 10)
    MineSup = CInt(CInt(Me.nbCases.Caption) * 7 / 10)
    Me.pourcent.Caption = "Entre " & CStr(MineInf) & " et " & CStr(MineSup)
End Sub

Private Sub nbMine_AfterUpdate()
    With Me.nbMine
        If Not IsNumeric(.Value) Or .Value = "" Then
            MsgBox "La valeur saisie n'est pas un nombre.", vbOKOnly + vbExclamation, "Valeur incorrecte"
            .Value = CInt(CInt(Me.nbCases.Caption) / 5)
        ElseIf .Value < MineInf Or .Value > MineSup Then
            MsgBox "Nombre de mines incorrect.", vbOKOnly + vbExclamation, "Nombre de mines incorrect"
            .Value = CInt(CInt(Me.nbCases.Caption) / 5)
        End If
    End With
End Sub

Private Sub SpinMine_SpinDown()
    With Me.nbMines
        If .Value = MineInf Then
            MsgBox "Nombre de mines minimal atteint.", vbOKOnly + vbExclamation, "Minimum atteint"
        Else
            .Value = .Value - 1
        End If
    End With
End Sub

Private Sub SpinMine_SpinUp()
    With Me.nbMines
        If .Value = MineSup Then
            MsgBox "Nombre de mines maximal atteint.", vbOKOnly + vbExclamation, "Maximum atteint"
        Else
            .Value = .Value + 1
        End If
    End With
End Sub

Private Sub commencer_Click()
    Call Module1.creerGrille(Me.nbLignes.Value, Me.nbCol.Value, Me.nbMines.Value)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    MineInf = 10
    MineSup = 70
End Sub
