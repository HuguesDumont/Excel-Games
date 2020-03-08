VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Mastermind (Par Hugues DUMONT)"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pos As Integer

Private Sub C1_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C1.BackColor
    Me.Controls.Item("J" & pos).Tag = "C1"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C2_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C2.BackColor
    Me.Controls.Item("J" & pos).Tag = "C2"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C3_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C3.BackColor
    Me.Controls.Item("J" & pos).Tag = "C3"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C4_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C4.BackColor
    Me.Controls.Item("J" & pos).Tag = "C4"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C5_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C5.BackColor
    Me.Controls.Item("J" & pos).Tag = "C5"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C6_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C6.BackColor
    Me.Controls.Item("J" & pos).Tag = "C6"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C7_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C7.BackColor
    Me.Controls.Item("J" & pos).Tag = "C7"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub C8_Click()
    Me.Controls.Item("J" & pos).BackColor = Me.C8.BackColor
    Me.Controls.Item("J" & pos).Tag = "C8"
    If pos Mod 4 = 0 Then
        Call afficherPositions
    End If
    pos = pos + 1
End Sub

Private Sub newGame_Click()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim ctrl As Control
    
    Randomize
    
    pos = 1
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) <> "CommandButton" Then
            If InStr(ctrl.Name, "C") < 1 Then ctrl.BackColor = -2147483644
        End If
    Next ctrl
    
    For i = 1 To 4
        Me.Controls.Item("S" & i).Tag = "C" & Int(Rnd * 8) + 1
    Next i
End Sub

Private Sub afficherPositions()
    Dim i As Long, j As Long, cpt As Long, k As Long
    Dim sol(3) As Long, trait(3) As Long
    Dim trouve As Boolean
    
    cpt = 0
    With Me.Controls
        For i = 0 To 3
            sol(i) = -1
            trait(i) = -1
        Next i
        
        For i = 1 To 4
            If .Item("S" & i).Tag = .Item("J" & pos - 4 + i).Tag Then
                sol(cpt) = i
                trait(cpt) = i
                cpt = cpt + 1
                .Item("P" & pos - 4 + cpt).BackColor = RGB(255, 0, 0)
            End If
        Next i
        
        If cpt = 4 Then
            victoire.Show
            Exit Sub
        End If
        
        For i = pos - 3 To pos
            trouve = False
            For j = 0 To 3
                If i - pos + 4 = trait(j) Then trouve = True
            Next j
            
            If Not trouve Then
                For j = 1 To 4
                    If i - pos + 4 <> j Then
                        If .Item("S" & j).Tag = .Item("J" & i).Tag Then
                            trouve = False
                            For k = 0 To 3
                                If sol(k) = j Then trouve = True
                            Next k
                            
                            If Not trouve Then
                                sol(cpt) = j
                                cpt = cpt + 1
                                .Item("P" & pos - 4 + cpt).BackColor = RGB(255, 255, 255)
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
        Next i
        
        If pos = 40 Then perdu.Show
    End With
End Sub
