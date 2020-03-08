Attribute VB_Name = "Module1"
Option Explicit

Public Sub ecrire()
    Dim ctrl As Control
    Dim code As String
    Dim Module As Variant
    
    Set Module = ThisWorkbook.VBProject.VBComponents
    
    With Module("Module1").CodeModule
        For Each ctrl In UserForm1.Controls
            If ctrl.BackColor = RGB(128, 0, 0) Then
                code = "Private Sub " & ctrl.Name & "_Click()" & Chr(13) & Chr(13) & "End Sub"
                .InsertLines .CountOfLines + 1, code
            End If
        Next ctrl
    End With
End Sub
