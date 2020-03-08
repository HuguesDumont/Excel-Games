VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} endGame 
   Caption         =   "Fin de partie"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4020
   OleObjectBlob   =   "endGame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "endGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub newGame_Click()
    GlobalVars.li_CoupsBons = 0
    GlobalVars.li_NbCoups = 0
    Unload Me
    optionSelector.Show
End Sub

Private Sub Quitter_Click()
    Call UserForm_QueryClose(False, 0)
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim c As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Mise à jours des records indivuels
    Set ws = ThisWorkbook.Sheets("Individuels")
    c = Asc(GlobalVars.ls_Col) - 64
    
    With ws.Cells(GlobalVars.ll_PlayerLine, (c / 2) + 1)
        If ((.Value = "") Or (.Value < GlobalVars.li_CoupsBons)) Then .Value = GlobalVars.li_CoupsBons
    End With
    
    ' Mise à jours du top 10
    Set ws = ThisWorkbook.Sheets("Records")
    For i = 2 To 11
        If (ws.Cells(i, c).Value < GlobalVars.li_CoupsBons) Then
            j = 11
            While (j <> i)
                ws.Cells(j, c).Value = ws.Cells(j - 1, c).Value
                ws.Cells(j, c - 1).Value = ws.Cells(j - 1, c - 1).Value
                j = j - 1
            Wend
            ws.Cells(i, c).Value = GlobalVars.li_CoupsBons
            ws.Cells(i, c - 1).Value = GlobalVars.ls_PlayerName
            Exit For
        End If
    Next i
    
    ' Affichage des résultats de la partie
    Me.touches.Caption = CStr(GlobalVars.li_CoupsBons)
    Me.rates.Caption = CStr(GlobalVars.li_NbCoups - GlobalVars.li_CoupsBons)
    Me.CoupsTotaux.Caption = CStr(GlobalVars.li_NbCoups)
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
