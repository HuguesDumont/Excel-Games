VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const rootPath As String = "C:\Users\Hugues\eclipse-workspace\Scrabble\resources\letters\"

Sub writeLetters()
    Dim i As Integer, j As Integer
    Dim sh As Worksheet
    Dim str As String
    
    Set sh = ThisWorkbook.Sheets(1)
    
    With sh
        For i = 1 To .Cells(1, sh.Columns.Count).End(xlToLeft).Column Step 3
            Call createLettersFile(sh.Cells(1, i).Value)
            For j = 2 To .Cells(Rows.Count, i).End(xlUp).Row
                str = .Cells(j, i).Value & Chr(9) & .Cells(j, i + 1).Value & Chr(9) & .Cells(j, i + 2).Value
                Call writeLineToFile(.Cells(1, i).Value, str)
            Next j
        Next i
    End With
End Sub

Sub writeLineToFile(fileName As String, textLine As String)
    fileName = rootPath + LCase(fileName) + ".txt"
    Open fileName For Append As #1
    Print #1, textLine
    Close #1
End Sub

Private Sub createLettersFile(fileName As String)
    On Error GoTo err
    Dim oFSO As Scripting.FileSystemObject
    Dim oFile As Object
    
    'Instanciation du FSO
    Set oFSO = New Scripting.FileSystemObject
    'Créer le fichier
    Set oFile = oFSO.CreateTextFile(rootPath + LCase(fileName) + ".txt", True, False)
fin:
    Set oFile = Nothing
    Set oFSO = Nothing
    Exit Sub
        
err:
    Select Case err.Number
        Case 58: MsgBox "Le dossier existe déjà"
        Case 76: MsgBox "Chemin incorrect"
        Case Else: MsgBox "Erreur inconnue"
    End Select
    
    Resume fin
End Sub
