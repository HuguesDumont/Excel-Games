VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public gDepRecherche As String
Public gJuste As Long
Public gTotal As Long

'-----------------------------------------------------------
'Procédure appelée lors du click sur la carte
'-----------------------------------------------------------
Sub CarteFrance_Click()
    Dim jeu As Worksheet, liste As Worksheet
    Dim i As Integer
    Dim lDep As String
    Dim lDepName As String
    
    Set jeu = ThisWorkbook.Sheets("Jeu des départements")
    Set liste = ThisWorkbook.Sheets("Liste départements")
    
    lDep = Application.Caller
    lDepName = DepartementNom(lDep)
    gTotal = gTotal + 1
    
    With jeu
        .Cells(3, 3).Value = lDepName
        For i = 2 To liste.Cells(Rows.Count, 1).End(xlUp).Row
            If .Cells(3, 3).Value = liste.Cells(i, 2).Value Then
                .Cells(3, 4).Value = liste.Cells(i, 3).Value
                .Cells(3, 5).Value = liste.Cells(i, 4).Value
                .Cells(3, 6).Value = liste.Cells(i, 5).Value
                .Cells(3, 7).Value = liste.Cells(i, 6).Value
            End If
        Next i
        '.Range("DepChoix").Value = lDepName
        
        If lDep = gDepRecherche Then
            MsgBox "PARFAIT !", vbOKOnly + vbExclamation, "Trouvé"
            gJuste = gJuste + 1
            NextDep
            For i = 3 To 7
                jeu.Cells(3, i).Value = ""
            Next i
        Else
            MsgBox "PERDU !", vbOKOnly + vbExclamation, "Perdu"
        End If
        MAJScore
    End With
End Sub

Public Property Let DepartementRecherche(pDep As String)
    gDepRecherche = pDep
End Property

Public Property Get DepartementNom(pDep As String) As String
    DepartementNom = ThisWorkbook.Sheets("Liste départements").Columns(1).Find(pDep).Offset(0, 1).Value
End Property

Public Sub MAJScore()
    With ThisWorkbook.Sheets("Jeu des départements")
        .Range("Score").Value = Format(gJuste & " / " & gTotal, "@")
        If gTotal >= 50 Then
            MsgBox "Votre Score :" & .Range("Score").Value
            gTotal = 0
            gJuste = 0
            .Range("Score").Value = ""
        End If
    End With
End Sub

Public Sub NextDep()
    Dim jeu As Worksheet, liste As Worksheet
    Dim i As Integer, j As Integer
    
    Set jeu = ThisWorkbook.Sheets("Jeu des départements")
    Set liste = ThisWorkbook.Sheets("Liste départements")
    
    i = Int(Rnd() * 96) + 2
    
    gDepRecherche = liste.Cells(i, 1).Value
    jeu.Cells(2, 3).Value = DepartementNom(gDepRecherche)
    For j = 4 To 7
        jeu.Cells(2, j).Value = liste.Cells(i, j - 1).Value
    Next j
End Sub
