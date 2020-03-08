Attribute VB_Name = "Module1"
Option Explicit

Function Sans_Accent(str As String) As String
    Dim accent As String
    Dim lettre As String
    Dim i As Long
    
    accent = UCase("éèêëiîïàâùûô")
    lettre = UCase("eeeeiiiaauuo")
    Sans_Accent = UCase(str)
    For i = 1 To Len(accent)
        If InStr(Sans_Accent, Mid(accent, i, 1)) > 0 Then
            Sans_Accent = Replace(Sans_Accent, Mid(accent, i, 1), Mid(lettre, i, 1))
        End If
    Next i
End Function
