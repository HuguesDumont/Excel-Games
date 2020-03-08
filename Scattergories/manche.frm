VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} manche 
   Caption         =   "Le jeu du petit bac"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14865
   OleObjectBlob   =   "manche.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "manche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isInit As Boolean
Private termine As Boolean
Private debut As Single

Private Sub correction_Click()
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim ctrl As Control
    
    Set ws = ThisWorkbook.Sheets("BDD")
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" And ctrl.BackColor = RGB(255, 0, 0) Then
            Me.Controls.Item("cor" & ctrl.Name).Visible = True
            For j = 1 To ws.Cells(1, Columns.Count).End(xlToLeft).Column
                If InStr(ws.Cells(1, j).Value, ctrl.Name) > 0 Then
                    For i = 2 To ws.Cells(Rows.Count, j).End(xlUp).Row
                        If Left(ws.Cells(i, j).Value, 1) = Me.LETTRE.Caption Then
                            Me.Controls.Item("cor" & ctrl.Name).Caption = ws.Cells(i, j).Value
                            Exit For
                        End If
                    Next i
                    If i = ws.Cells(Rows.Count, j).End(xlUp).Row + 1 Then Me.Controls.Item("cor" & ctrl.Name).Caption = "PAS DE REPONSE"
                    Exit For
                End If
            Next j
        End If
    Next ctrl
    
    Me.correction.Enabled = False
End Sub

Private Sub lettreAlea_Click()
    Dim Cible As Byte
    Dim ctrl As Control
    Dim prev As String
 
    Randomize
    
    prev = Me.LETTRE.Caption
    
    While prev = Me.LETTRE.Caption
        Cible = Int((26 * Rnd) + 1)
        Me.LETTRE.Caption = Chr(Cible + 64)
        Me.LabelPOINTS.Caption = 0
    Wend
    
    Me.lettreAlea.Enabled = False
    Me.correction.Enabled = False
    Me.valider.Enabled = True
    termine = False
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Enabled = True
            ctrl.BackColor = RGB(255, 255, 255)
            ctrl.Value = ""
            With Me.Controls.Item("cor" & ctrl.Name)
                .Caption = ""
                .Visible = False
            End With
        End If
    Next ctrl
    
    isInit = False
    debut = Now
    Call UserForm_Activate
End Sub

Private Sub UserForm_Activate()
    Dim debut As Single
    Dim i As Long
    Dim ctrl As Control
    
    If isInit Then Exit Sub
    
    While ((Format(Now - debut, "ss") < 60) And (Not termine))
        DoEvents
    Wend
    
    If Not termine Then MsgBox "Fin du temps, les points vont être calculés", vbOKOnly + vbInformation, "Temps écoulé"
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl.Name) = "TextBox" Then ctrl.Enabled = False
    Next ctrl
    
    Call valider_Click
    
    Me.lettreAlea.Enabled = True
    Me.correction.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    Dim ctrl As Control
    Dim i As Integer
    
    isInit = True
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) <> "TextBox" And TypeName(ctrl) <> "CommandButton" Then ctrl.BackColor = Me.LETTRE.BackColor
        If InStr(ctrl.Name, "cor") > 0 And ctrl.Name <> "correction" Then
            ctrl.BackColor = RGB(100, 100, 255)
            ctrl.Visible = False
        End If
        If TypeName(ctrl) = "TextBox" Then
            For i = 0 To UBound(Module1.categ, 2)
                If Module1.categ(0, i) = ctrl.Name Then
                    If Module1.categ(1, i) = "TRUE" Then
                        ctrl.Enabled = True
                    Else
                        ctrl.Enabled = False
                    End If
                    Exit For
                End If
            Next i
        End If
    Next ctrl
    
    termine = False
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    termine = True
    ThisWorkbook.Close
    Application.Visible = True
End Sub

Private Sub valider_Click()
    Dim i As Long, j As Long
    Dim ctrl As Control
    Dim ws As Worksheet
    Dim trouve As Boolean
    Dim cpt As Integer
    Dim rng As Range
    
    If termine = False Then termine = True
    
    Set ws = ThisWorkbook.Sheets("BDD")
    cpt = 0
    Call Module1.initTab
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = Module1.majacc(ctrl.Value)
            trouve = False
            For j = 1 To ws.Cells(1, Columns.Count).End(xlToLeft).Column
                If InStr(ws.Cells(1, j).Value, ctrl.Name) > 0 Then
                    If UCase(Left(ctrl.Value, 1)) = Me.LETTRE.Caption Then
                        Set rng = ws.Columns(j).Find(UCase(ctrl.Value))
                        
                        If Not rng Is Nothing Then
                            ctrl.BackColor = RGB(0, 255, 0)
                            cpt = cpt + 1
                        Else
                            ctrl.BackColor = RGB(255, 0, 0)
                        End If
                    Else
                        ctrl.BackColor = RGB(255, 0, 0)
                    End If
                    
                    Exit For
                End If
            Next j
        End If
    Next ctrl
    
    Me.LabelPOINTS.Caption = cpt
    Me.valider.Enabled = False
    Me.lettreAlea.Enabled = True
    isInit = True
End Sub


