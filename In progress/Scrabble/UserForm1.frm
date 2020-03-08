VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Scrabble (created par Hugues DUMONT)"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15345
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ComboBox1_Change()
    Dim src As String
    Dim i As Long, col As Long, ligneAjout As Long
    Dim j As Integer
    Dim plage As Range
    
    Variables.numLangue = Me.ComboBox1.ListIndex + 1
    Variables.langue = Variables.fDic.Cells(2, Variables.numLangue).Value
    Me.Label3.Caption = Variables.fDic.Cells(4, Variables.numLangue).Value
    src = Variables.srcImg & Variables.langue & ".bmp"
    
    Set Me.Image1.Picture = LoadPicture(src)
    Me.uneLettre.Caption = Variables.fDic.Cells(5, Variables.numLangue).Value
    Me.lesPoints.Caption = Variables.fDic.Cells(6, Variables.numLangue).Value
    Me.leNombre.Caption = Variables.fDic.Cells(7, Variables.numLangue).Value
    Me.startGame.Caption = Variables.fDic.Cells(8, Variables.numLangue).Value
    Me.ResetGame.Caption = Variables.fDic.Cells(9, Variables.numLangue).Value
    Me.NextPlayer.Caption = Variables.fDic.Cells(10, Variables.numLangue).Value
    Me.addWord.Caption = Variables.fDic.Cells(12, Variables.numLangue).Value
    
    Fonctions.initGame Variables.numLangue * 3 - 2, 3
    
    Variables.txtNextJ = Me.NextPlayer.Caption
    Variables.infoTour = Variables.fDic.Cells(11, Variables.numLangue).Value
    
    Set plage = Variables.fPions.Range(Variables.fPions.Cells(4, Variables.numLangue * 3 - 2), _
        Variables.fPions.Cells(Variables.fPions.Cells(4, Variables.numLangue * 3 - 2).End(xlDown).Row, _
        Variables.numLangue * 3))
        
    Me.ListBox1.RowSource = plage.Address
    
End Sub

Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = Fonctions.scrollFleche(Me, Me.ComboBox1.Name, KeyCode)
End Sub

Private Sub L1_Click()
    
End Sub

Private Sub L2_Click()
    
End Sub

Private Sub L3_Click()
    
End Sub

Private Sub L4_Click()
    
End Sub

Private Sub L5_Click()
    
End Sub

Private Sub L6_Click()
    
End Sub

Private Sub L7_Click()
    
End Sub

Private Sub startGame_Click()
    Dim i As Integer
    
    Me.SpinButton1.Enabled = False
    Me.ComboBox1.Enabled = False
    Me.startGame.Enabled = False
    Me.NextPlayer.Enabled = True
    Me.ResetGame.Enabled = True
    
    Variables.fGame.Cells(1, 2).Value = Variables.numLangue
    
    For i = 1 To CInt(Me.Label4.Caption)
        Fonctions.distrib (2 * i + 1)
    Next i
    
    Call initial
    
    Variables.player = 1
    
    Call displayLetters(1)
End Sub

Private Sub NextPlayer_Click()
    ''''''Gérer la validation du mot, le calcul des points et la distribution des nouvelles lettres
    
    ''''''

    ''''''On passe au joueur suivant
    If Variables.player = CInt(Me.Label4.Caption) Then
        Variables.player = 1
    Else
        Variables.player = Variables.player + 1
    End If
    Application.Wait Time + TimeSerial(0, 0, 3)
    
    Call displayLetters(Variables.player)
End Sub

Private Sub ResetGame_Click()
    Dim i As Integer
    
    Me.SpinButton1.Enabled = True
    Me.ComboBox1.Enabled = True
    Me.startGame.Enabled = True
    Me.NextPlayer.Enabled = False
    Me.ResetGame.Enabled = False
    
    For i = 1 To 7
        Me.Controls.Item("L" & i).Caption = ""
    Next i
    
    Variables.player = 1
    Call initial
End Sub

Private Sub SpinButton1_SpinDown()
    If CInt(Me.Label4.Caption) <> 1 Then Me.Label4.Caption = CInt(Me.Label4.Caption) - 1
End Sub

Private Sub SpinButton1_SpinUp()
    If CInt(Me.Label4.Caption) <> 4 Then Me.Label4.Caption = CInt(Me.Label4.Caption) + 1
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    
    Variables.srcImg = ThisWorkbook.Path & "\"

    Variables.langue = "Français"
    Variables.numLangue = 15
    
    Set Variables.fDic = ThisWorkbook.Sheets("Langues")
    Set Variables.fPions = ThisWorkbook.Sheets("Pions")
    Set Variables.fGame = ThisWorkbook.Sheets("Jeu")
    Variables.player = 1

    Me.Label3.Caption = fDic.Cells(4, 15).Value
    Me.Label4.Caption = 1
    
    Me.startGame.Enabled = True
    Me.SpinButton1.Enabled = True
    Me.ComboBox1.Enabled = True
    Me.NextPlayer.Enabled = False
    Me.ResetGame.Enabled = False
    
    Me.ComboBox1.Clear
    For i = 1 To 38
        Me.ComboBox1.AddItem Variables.fDic.Cells(3, i).Value
    Next i
    Me.ComboBox1.ListIndex = 14
    
    Fonctions.initGame Variables.numLangue * 3 - 2, 3
    Call initial
End Sub

Private Sub displayLetters(j As Integer)
    Dim i As Integer
    
    With Variables.fGame
        For i = 1 To 7
            Me.Controls.Item("L" & CStr(i)).Caption = .Cells(2 * j + 1, 4 + i).Value
        Next i
    End With
End Sub

Private Sub initial()
    Dim i As Long, nom As String
    
    nom = "TextBox"
    
    ''''''initialisation
    For i = 1 To 225
        With Me.Controls.Item("TextBox" & i)
            .Tag = 1
            .ControlTipText = 1
        End With
    Next i
    ''''''
    
    ''''''mots triples
    For i = 1 To 225 Step 7
        Me.Controls.Item("TextBox" & i).ControlTipText = 3
    Next i
    ''''''
    
    With Me.Controls
    
        ''''''mots doubles
        .Item(nom & 17).ControlTipText = 2
        .Item(nom & 29).ControlTipText = 2
        .Item(nom & 33).ControlTipText = 2
        .Item(nom & 43).ControlTipText = 2
        .Item(nom & 49).ControlTipText = 2
        .Item(nom & 57).ControlTipText = 2
        .Item(nom & 65).ControlTipText = 2
        .Item(nom & 71).ControlTipText = 2
        .Item(nom & 113).ControlTipText = 2
        .Item(nom & 161).ControlTipText = 2
        .Item(nom & 169).ControlTipText = 2
        .Item(nom & 177).ControlTipText = 2
        .Item(nom & 183).ControlTipText = 2
        .Item(nom & 193).ControlTipText = 2
        .Item(nom & 197).ControlTipText = 2
        .Item(nom & 207).ControlTipText = 2
        ''''''
        
        ''''''Lettres doubles
        .Item(nom & 4).Tag = 2
        .Item(nom & 12).Tag = 2
        .Item(nom & 37).Tag = 2
        .Item(nom & 39).Tag = 2
        .Item(nom & 46).Tag = 2
        .Item(nom & 53).Tag = 2
        .Item(nom & 60).Tag = 2
        .Item(nom & 93).Tag = 2
        .Item(nom & 97).Tag = 2
        .Item(nom & 99).Tag = 2
        .Item(nom & 103).Tag = 2
        .Item(nom & 109).Tag = 2
        .Item(nom & 117).Tag = 2
        .Item(nom & 123).Tag = 2
        .Item(nom & 129).Tag = 2
        .Item(nom & 133).Tag = 2
        .Item(nom & 166).Tag = 2
        .Item(nom & 173).Tag = 2
        .Item(nom & 180).Tag = 2
        .Item(nom & 187).Tag = 2
        .Item(nom & 189).Tag = 2
        .Item(nom & 214).Tag = 2
        .Item(nom & 222).Tag = 2
        ''''''
        
        ''''''lettres triples
        .Item(nom & 21).Tag = 3
        .Item(nom & 25).Tag = 3
        .Item(nom & 77).Tag = 3
        .Item(nom & 81).Tag = 3
        .Item(nom & 85).Tag = 3
        .Item(nom & 89).Tag = 3
        .Item(nom & 137).Tag = 3
        .Item(nom & 141).Tag = 3
        .Item(nom & 145).Tag = 3
        .Item(nom & 149).Tag = 3
        .Item(nom & 201).Tag = 3
        .Item(nom & 205).Tag = 3
        ''''''
    End With
End Sub
