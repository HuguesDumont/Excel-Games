VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pendu 
   Caption         =   "Jeu du pendu (par Hugues DUMONT)"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10650
   OleObjectBlob   =   "pendu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pendu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numImg As String
Private endGame As Boolean

Private Sub affExcel_Click()
    Application.Visible = True
    Unload Me
End Sub

Private Sub langues_Change()
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("param")
    
    Me.newGame.Caption = ws.Cells(2, Me.langues.ListIndex + 1).Value
    Me.languages.Caption = ws.Cells(3, Me.langues.ListIndex + 1).Value
    Me.sol.Caption = ws.Cells(4, Me.langues.ListIndex + 1).Value
    Me.letter.Caption = ws.Cells(5, Me.langues.ListIndex + 1).Value
    Me.wording.Caption = ws.Cells(6, Me.langues.ListIndex + 1).Value
    Me.vali.Caption = ws.Cells(7, Me.langues.ListIndex + 1).Value
    Me.tentes.Caption = ws.Cells(8, Me.langues.ListIndex + 1).Value
    Call newGame_Click
End Sub

Private Sub lettre_Change()
    Dim i As Long
    Dim trouve As Boolean
    Dim srcimg As String
    Const alpha As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    Me.lettre.Value = UCase(Module1.Sans_Accent(Me.lettre.Value))
    
    If InStr(alpha, Me.lettre.Value) < 1 Or Me.lettre.Value = "" Or endGame Then
        Me.lettre.Value = ""
        Exit Sub
    End If
    
    trouve = False
    For i = 1 To Len(Me.solution.Caption)
        If Mid(Me.solution.Caption, i, 1) = Me.lettre.Value Then
            trouve = True
            Me.masque.Caption = Left(Me.masque.Caption, i - 1) & Me.lettre.Value & Right(Me.masque.Caption, Len(Me.solution.Caption) - i)
        End If
    Next i
    If InStr(Me.tried.Caption, Me.lettre.Value) < 1 Then Me.tried.Caption = Me.tried.Caption & " " & Me.lettre.Value
    
    Call tried_Click
    Me.lettre.Value = ""

    If Not trouve Then
        numImg = numImg + 1
        srcimg = ThisWorkbook.Path & "\Pendu" & CStr(numImg) & ".bmp"
        Me.Image.Picture = LoadPicture(srcimg)
        If numImg = 10 Then
            endGame = True
            Me.masque.ForeColor = RGB(255, 0, 0)
            Me.solution.Visible = True
            Me.solution.ForeColor = RGB(0, 255, 0)
            sons.jouerSon ("lost")
        End If
    Else
        Call masque_Click
    End If
End Sub

Private Sub masque_Click()
    If Me.masque.Caption = Me.solution.Caption Then
        endGame = True
        sons.jouerSon ("Victoire")
        victoire.Show
        Call newGame_Click
    End If
End Sub

Private Sub newGame_Click()
    Dim leMot As String, srcimg
    
    srcimg = ThisWorkbook.Path & "\Pendu0.bmp"
    numImg = 0
    endGame = False
    
    Me.masque.ForeColor = RGB(0, 0, 0)
    leMot = tirerMot(Me.langues.ListIndex + 1)
    Me.solution.Caption = formaterMot(leMot)
    Me.masque.Caption = masquerMot(Me.solution.Caption)
    Me.tried.Caption = ""
    Me.Image.Picture = LoadPicture(srcimg)
    Me.mot.Value = ""
    Me.solution.Visible = False
End Sub

Private Sub tried_Click()
    Dim tabmot() As String, tmp As String
    Dim i As Long, j As Long
    
    If Len(Me.tried.Caption) < 2 Then Exit Sub
    
    tabmot = Split(Me.tried.Caption)
    For i = 0 To UBound(tabmot)
        tabmot(i) = Trim(tabmot(i))
    Next i
    
    For i = 0 To UBound(tabmot) - 1
        For j = i + 1 To UBound(tabmot)
            If tabmot(j) < tabmot(i) Then
                tmp = tabmot(i)
                tabmot(i) = tabmot(j)
                tabmot(j) = tmp
            End If
        Next j
    Next i
    
    Me.tried.Caption = tabmot(0)
    For i = 1 To UBound(tabmot)
        Me.tried.Caption = Me.tried.Caption & " " & tabmot(i)
    Next i
    Call lesMenus.AddMinimiseButton
End Sub

Private Sub UserForm_Activate()
    Call lesMenus.AddMinimiseButton
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim ws As Worksheet
    Dim leMot As String, srcimg As String
    
    srcimg = ThisWorkbook.Path & "\Pendu0.bmp"
    Set ws = ThisWorkbook.Sheets("dic")

    Me.langues.Clear
    For i = 1 To ws.Cells(1, Columns.Count).End(xlToLeft).Column
        Me.langues.AddItem ws.Cells(1, i).Value
    Next i
    
    Me.langues.ListIndex = 0
    leMot = tirerMot(1)
    Me.solution.Caption = formaterMot(leMot)
    Me.masque.Caption = masquerMot(Me.solution.Caption)
    Me.tried.Caption = ""
    Me.Image.Picture = LoadPicture(srcimg)
    Me.solution.Visible = False
    Me.masque.ForeColor = RGB(0, 0, 0)
    
    endGame = False
    numImg = 0
End Sub

Private Function formaterMot(leMot As String)
    Dim tabmot() As String
    Dim i As Long
    On Error Resume Next
    
    ReDim tabmot(Len(leMot) - 1)
    For i = 1 To Len(leMot)
        tabmot(i - 1) = Mid(leMot, i, 1)
        If tabmot(i - 1) <> " " And tabmot(i - 1) <> "_" Then
            tabmot(i - 1) = tabmot(i - 1) & " "
            If i = Len(leMot) Then tabmot(i - 1) = Trim(tabmot(i - 1))
        End If
    Next i
    
    formaterMot = ""
    For i = 1 To Len(leMot)
        formaterMot = formaterMot & tabmot(i - 1)
    Next i
End Function

Private Function tirerMot(i) As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("dic")
    Randomize
    tirerMot = ws.Cells(Int((ws.Cells(Rows.Count, i).End(xlUp).Row * Rnd) + 2), i)
End Function

Private Function masquerMot(mot As String) As String
    Dim i As Long
    
    i = 1
    While i <= Len(mot)
        If ((Mid(mot, i, 1) <> "-") And (Mid(mot, i, 1) <> " ") And (Mid(mot, i, 1) <> "_")) Then
            mot = Replace(mot, Mid(mot, i, 1), "_")
        End If
        i = i + 1
    Wend
    masquerMot = mot
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then
        ThisWorkbook.Close
        Application.Quit
    End If
End Sub

Private Sub vali_Click()
    Dim i As Long
    Dim leMot As String, srcimg As String
    
    
    Const alpha As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ-"
    
    If endGame Then Exit Sub
    
    Me.mot.Value = UCase(Module1.Sans_Accent(Me.mot.Value))
    
    For i = 1 To Len(Me.mot.Value)
        If InStr(alpha, Mid(Me.mot.Value, i, 1)) < 1 Then
            MsgBox "Il y a des caractères incorrects dans ce mot. Cette tentative ne compte donc pas.", vbOKOnly + vbExclamation, "Mot invalide"
            Exit Sub
        End If
    Next i
    
    If formaterMot(Me.mot.Value) = Me.solution.Caption Then
        endGame = True
        sons.jouerSon ("Victoire")
        victoire.Show
        Call newGame_Click
    Else
        numImg = numImg + 1
        srcimg = ThisWorkbook.Path & "\Pendu" & CStr(numImg) & ".bmp"
        Me.Image.Picture = LoadPicture(srcimg)
        If numImg = 10 Then
            endGame = True
            Me.masque.ForeColor = RGB(255, 0, 0)
            Me.solution.Visible = True
            Me.solution.ForeColor = RGB(0, 255, 0)
            sons.jouerSon ("lost")
        End If
    End If
        
    leMot = formaterMot(Me.mot.Value)
End Sub
