VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} perdu 
   Caption         =   "PERDU :("
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   OleObjectBlob   =   "perdu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "perdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub quitter_Click()
    Unload Me
    ThisWorkbook.Close
    Application.Quit
End Sub

Private Sub recommencer_Click()
    Unload Me
    param.Show
End Sub

Private Sub solution_Click()
    Dim ws As Worksheet
    Dim c As Range, rng As Range
    
    Set ws = ThisWorkbook.Sheets("Valeurs")
    Set rng = ThisWorkbook.Sheets("Démineur").Range("B2:" & ws.Cells(1, 65).Value & ws.Cells(1, 64).Value)
    
    Me.Left = 800
    
    For Each c In rng
        c.Value = ws.Cells(c.Row, c.Column).Value
        If c.Value <> "X" Then c.Font.Color = ws.Cells(c.Row, c.Column).Font.Color
    Next c
End Sub
