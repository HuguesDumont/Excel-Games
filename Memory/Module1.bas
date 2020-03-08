Attribute VB_Name = "Module1"
Option Explicit

Public previous As Integer, trouvees As Integer, nbCoups As Integer
Public ima(9) As String, plateau(19) As Integer

Public Sub jouer(val As Integer)
    Dim i As Integer
    With Memory.Controls
        If previous = -1 Then
            previous = val
            .Item("Image" & val + 1).Picture = LoadPicture(ima(plateau(val)))
            .Item("Image" & val + 1).Enabled = False
        Else
            .Item("Image" & val + 1).Picture = LoadPicture(ima(plateau(val)))
            Application.Wait Time + TimeSerial(0, 0, 2)
            If plateau(previous) = plateau(val) Then
                .Item("Image" & val + 1).Enabled = False
                .Item("Image" & previous + 1).Enabled = False
                .Item("Image" & previous + 1).Picture = Nothing
                .Item("Image" & val + 1).Picture = Nothing
                .Item("Image" & previous + 1).BackColor = RGB(255, 255, 255)
                .Item("Image" & val + 1).BackColor = RGB(255, 255, 255)
                trouvees = trouvees + 1
                If trouvees = 10 Then
                    nbCoups = nbCoups + 1
                    For i = 0 To 19
                        .Item("Image" & i + 1).Picture = LoadPicture(ima(plateau(i)))
                    Next i
                    MsgBox ("Félicitations, tu as gagné!" & Chr(13) & "Victoire en " & nbCoups & " coups")
                    nbCoups = nbCoups - 1
                End If
            Else
                .Item("Image" & val + 1).Picture = Nothing
                .Item("Image" & previous + 1).Picture = Nothing
                .Item("Image" & val + 1).Enabled = True
                .Item("Image" & previous + 1).Enabled = True
            End If
            previous = -1
            nbCoups = nbCoups + 1
        End If
        .Item("Label1").Caption = "Coups joués : " & Module1.nbCoups
    End With
End Sub
