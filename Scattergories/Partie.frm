VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Partie 
   Caption         =   "Paramètres de partie"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   OleObjectBlob   =   "Partie.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Partie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub annuler_Click()
    Unload Me
    ThisWorkbook.Close
    Application.Visible = True
End Sub

Private Sub SpinButton1_SpinDown()
    If IsNumeric(Me.nbJoueurs.Value) And Not _
        (Me.nbJoueurs.Value <= 1 Or InStr(CStr(Me.nbJoueurs.Value), ".") > 0 Or InStr(CStr(Me.nbJoueurs.Value), ",") > 0) Then
        
        Me.nbJoueurs.Value = Me.nbJoueurs.Value - 1
    End If
End Sub

Private Sub SpinButton1_SpinUp()
    If IsNumeric(Me.nbJoueurs.Value) And Not (InStr(CStr(Me.nbJoueurs.Value), ".") > 0 Or InStr(CStr(Me.nbJoueurs.Value), ",") > 0) Then
        Me.nbJoueurs.Value = Me.nbJoueurs.Value + 1
    End If
End Sub

Private Sub valider_Click()
    Dim ctrl As Control
    Dim cpt As Integer, nbOK As Integer
    
    If Not IsNumeric(Me.nbJoueurs.Value) Or Me.nbJoueurs.Value < 1 Or InStr(CStr(Me.nbJoueurs.Value), ".") > 0 Or InStr(CStr(Me.nbJoueurs.Value), ",") > 0 Then
        MsgBox "Le nombre de joueurs n'est pas valide. Veuillez le corriger", vbExclamation + vbOKOnly, "Données incorrectes"
        Exit Sub
    End If
    
    cpt = 0
    nbOK = 0
    For Each ctrl In Me.cat.Controls
        Module1.categ(0, cpt) = CStr(ctrl.Tag)
        If ctrl.Value Then
            Module1.categ(1, cpt) = "TRUE"
        Else
            Module1.categ(1, cpt) = "FALSE"
        End If
        cpt = cpt + 1
        If ctrl.Value Then nbOK = nbOK + 1
    Next ctrl
    
    If nbOK < 1 Then
        MsgBox "Vous devez sélectionner au moins une catégorie", vbOKOnly + vbExclamation, "Aucune catégorie sélectionnée"
        Exit Sub
    End If
    
    Unload Me
    manche.Show
End Sub
