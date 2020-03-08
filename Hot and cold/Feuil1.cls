VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo fin
    If Target.Cells.Count <> 1 Then Exit Sub
    If Not Module1.premier Then
        Call Module1.verif
    Else
        Module1.premier = False
    End If
fin:
End Sub
