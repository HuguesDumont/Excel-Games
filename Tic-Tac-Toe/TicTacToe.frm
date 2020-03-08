VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TicTacToe 
   Caption         =   "Tic-Tac-Toe by Hugues DUMONT"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   OleObjectBlob   =   "TicTacToe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lb_Player As Boolean ' Player 1 = True, Player 2 = False
Private lb_Finished As Boolean ' Game finished = True

Private Sub NewGame_Click()
    Call UserForm_Initialize
End Sub

Private Sub ExitGame_Click()
    Call UserForm_QueryClose(False, 0)
End Sub

Private Sub UserForm_Initialize()
    Dim lo_Ctrl As MSForms.Control
    
    lb_Player = True
    lb_Finished = False
    
    For Each lo_Ctrl In Me.Controls
        If (Len(lo_Ctrl.Name) = 2) Then
            lo_Ctrl.Caption = ""
            lo_Ctrl.ForeColor = RGB(0, 0, 0)
            lo_Ctrl.BackColor = &H8000000F
        End If
    Next lo_Ctrl
    
    Me.LabelP1.BackColor = &H8000000D
    Me.LabelP2.BackColor = &H8000000F
End Sub

Private Sub BC_Click()
    Call sb_Play(Me.BC)
End Sub

Private Sub BL_Click()
    Call sb_Play(Me.BL)
End Sub

Private Sub BR_Click()
    Call sb_Play(Me.BR)
End Sub

Private Sub MC_Click()
    Call sb_Play(Me.MC)
End Sub

Private Sub ML_Click()
    Call sb_Play(Me.ML)
End Sub

Private Sub MR_Click()
    Call sb_Play(Me.MR)
End Sub

Private Sub TC_Click()
    Call sb_Play(Me.TC)
End Sub

Private Sub TL_Click()
    Call sb_Play(Me.TL)
End Sub

Private Sub TR_Click()
    Call sb_Play(Me.TR)
End Sub

Private Sub sb_Play(ByRef po_Label As MSForms.Label)
    If ((Not lb_Finished) And (po_Label.Caption = "")) Then
        If (lb_Player) Then
            po_Label.Caption = "O"
            po_Label.ForeColor = RGB(0, 255, 0)
        Else
            po_Label.Caption = "X"
            po_Label.ForeColor = RGB(255, 0, 0)
        End If
        
        Call sb_CheckWin
        lb_Player = (Not lb_Player)
        
        If (lb_Player) Then
            Me.LabelP1.BackColor = &H8000000D
            Me.LabelP2.BackColor = &H8000000F
        Else
            Me.LabelP1.BackColor = &H8000000F
            Me.LabelP2.BackColor = &H8000000D
        End If
    End If
End Sub

Private Sub sb_CheckWin()
    Dim lo_Ctrl As Control
    Dim ls_Message As String
    Dim lb_Equal As Boolean
    Dim ls_WinPlay As String
    
    ls_Message = ""
    ls_WinPlay = IIf(lb_Player, "Player 1 ", "Player 2 ") & "win with "
    lb_Equal = False
    
    If ((Me.TC.Caption = Me.TL.Caption) And (Me.TR.Caption = Me.TC.Caption) And (Me.TC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " top line !"
        Call sb_HighlightWin(Me.TC, Me.TL, Me.TR)
    ElseIf ((Me.MC.Caption = Me.ML.Caption) And (Me.MR.Caption = Me.MC.Caption) And (Me.MC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " middle line !"
        Call sb_HighlightWin(Me.MC, Me.ML, Me.MR)
    ElseIf ((Me.BC.Caption = Me.BL.Caption) And (Me.BR.Caption = Me.BC.Caption) And (Me.BC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " bottom line !"
        Call sb_HighlightWin(Me.BC, Me.BL, Me.BR)
    ElseIf ((Me.TL.Caption = Me.ML.Caption) And (Me.ML.Caption = Me.BL.Caption) And (Me.TL.Caption <> "")) Then
        ls_Message = ls_WinPlay & " left column !"
        Call sb_HighlightWin(Me.TL, Me.ML, Me.BL)
    ElseIf ((Me.TC.Caption = Me.MC.Caption) And (Me.MC.Caption = Me.BC.Caption) And (Me.TC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " center column !"
        Call sb_HighlightWin(Me.TC, Me.MC, Me.BC)
    ElseIf ((Me.TR.Caption = Me.MR.Caption) And (Me.MR.Caption = Me.BR.Caption) And (Me.TR.Caption <> "")) Then
        ls_Message = ls_WinPlay & " right column !"
        Call sb_HighlightWin(Me.TR, Me.MR, Me.BR)
    ElseIf ((Me.TL.Caption = Me.MC.Caption) And (Me.BR.Caption = Me.MC.Caption) And (Me.MC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " diagonal top left to bottom right !"
        Call sb_HighlightWin(Me.TL, Me.MC, Me.BR)
    ElseIf ((Me.TR.Caption = Me.MC.Caption) And (Me.MC.Caption = Me.BL.Caption) And (Me.MC.Caption <> "")) Then
        ls_Message = ls_WinPlay & " diagonal bottom left to top right !"
        Call sb_HighlightWin(Me.BR, Me.MC, Me.TL)
    Else
        lb_Equal = True
        For Each lo_Ctrl In Me.Controls
            If (Len(lo_Ctrl.Name) = 2) Then
                If (lo_Ctrl.Caption = "") Then
                    lb_Equal = False
                    Exit For
                End If
            End If
        Next lo_Ctrl
        
        If (lb_Equal) Then
            ls_Message = "No winner for this game !"
        End If
    End If
    
    If (ls_Message <> "") Then
        lb_Finished = True
        
        If (lb_Equal) Then
            Me.NoWin.Caption = CInt(Me.NoWin.Caption) + 1
        ElseIf (lb_Player) Then
            Me.P1Win.Caption = CInt(Me.P1Win.Caption) + 1
        Else
            Me.P2Win.Caption = CInt(Me.P2Win.Caption) + 1
        End If
        
        MsgBox ls_Message
    End If
End Sub

Private Sub sb_HighlightWin(ByRef po_First As MSForms.Label, ByRef po_Second As MSForms.Label, ByRef po_third As MSForms.Label)
    po_First.BackColor = RGB(173, 255, 47)
    po_Second.BackColor = RGB(173, 255, 47)
    po_third.BackColor = RGB(173, 255, 47)
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
