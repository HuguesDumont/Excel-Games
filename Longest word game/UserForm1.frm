VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Le mot le plus long (par Hugues DUMONT)"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fini As Boolean

Private Sub word_Change()
    Me.word.Value = UCase(Me.word.Value)
End Sub

Private Sub voyelle_Click()
    Const voy As String = "AEIOUY"
    
    Randomize
    
    Me.tirage.Caption = Trim(Me.tirage.Caption & " " & Mid(voy, Int(Rnd * 6) + 1, 1))
    
    If Len(Me.tirage.Caption) = 19 Then
        Me.valider.Enabled = True
        Me.consonne.Enabled = False
        Me.voyelle.Enabled = False
        
        start = Timer
        fini = False
        
        Me.longMax.Caption = Module1.lePlusLong(Replace(Me.tirage.Caption, " ", ""))
        
        fini = True
        While Timer - start < 30
            Me.Label1.Caption = (30 - Round(Timer - start, 0)) & " s"
            DoEvents
        Wend
        
        MsgBox "La recherche des meilleures solutions est terminée.", vbOKOnly + vbInformation, "Recherche de solutions terminée"
        
        Call valider_Click
        
    End If
End Sub

Private Sub consonne_Click()
    Const con As String = "BCDFGHJKLMNPQRSTVWXZ"
    
    Randomize
    
    Me.tirage.Caption = Trim(Me.tirage.Caption & " " & Mid(con, Int(Rnd * 20) + 1, 1))
    
    If Len(Me.tirage.Caption) = 19 Then
        Me.valider.Enabled = True
        Me.consonne.Enabled = False
        Me.voyelle.Enabled = False
        
        start = Timer
        fini = False
        
        Me.longMax.Caption = Module1.lePlusLong(Replace(Me.tirage.Caption, " ", ""))
        
        fini = True
        While Timer - start < 30
            Me.Label1.Caption = (30 - Round(Timer - start, 0)) & " s"
            DoEvents
        Wend
        
        Call valider_Click
    End If
End Sub

Private Sub exitGame_Click()
    Unload Me
    ThisWorkbook.Saved = True
    Application.Quit
End Sub

Private Sub newGame_Click()
    Call UserForm_Initialize
End Sub

Private Sub UserForm_Initialize()
    Me.voyelle.Enabled = True
    Me.consonne.Enabled = True
    Me.valider.Enabled = False
    
    Me.ListBox1.Clear
    Me.word.Value = ""
    
    Me.tirage.Caption = ""
    Me.Height = 210
End Sub

Private Sub valider_Click()
    If Not fini Then MsgBox "La recherche de solutions n'est peut-être pas terminée." & Chr(13) & "Pour cela il faut attendre la fin du compte à rebours.", vbOKOnly + vbInformation, "Solutions potentiellement incomplètes"
    
    Me.Height = 433
    Me.Top = 100
End Sub
