Attribute VB_Name = "Module1"
Option Explicit

Public tableau() As String
Public categ(1, 9) As String

Public Function majacc(txt As String) As String
    Dim i As Long

    txt = UCase(txt)
    For i = 0 To 24
        txt = Replace(txt, tableau(0, i), tableau(1, i))
    Next i
    majacc = txt
End Function

Public Sub initTab()
    Dim i As Long
    
    ReDim tableau(1, 25)
    
    tableau(0, 0) = "À"
    tableau(0, 1) = "Á"
    tableau(0, 2) = "Â"
    tableau(0, 3) = "Ã"
    tableau(0, 4) = "Ä"
    tableau(0, 5) = "Ç"
    tableau(0, 6) = "É"
    tableau(0, 7) = "È"
    tableau(0, 8) = "Ê"
    tableau(0, 9) = "Ë"
    tableau(0, 10) = "Ì"
    tableau(0, 11) = "Í"
    tableau(0, 12) = "Î"
    tableau(0, 13) = "Ï"
    tableau(0, 14) = "Ò"
    tableau(0, 15) = "Ó"
    tableau(0, 16) = "Ô"
    tableau(0, 17) = "Õ"
    tableau(0, 18) = "Ö"
    tableau(0, 19) = "Ù"
    tableau(0, 20) = "Ú"
    tableau(0, 21) = "Û"
    tableau(0, 22) = "Ü"
    tableau(0, 23) = "Ý"
    tableau(0, 24) = "Ÿ"
    
    tableau(1, 0) = "A"
    tableau(1, 1) = "A"
    tableau(1, 2) = "A"
    tableau(1, 3) = "A"
    tableau(1, 4) = "A"
    tableau(1, 5) = "C"
    tableau(1, 6) = "E"
    tableau(1, 7) = "E"
    tableau(1, 8) = "E"
    tableau(1, 9) = "E"
    tableau(1, 10) = "I"
    tableau(1, 11) = "I"
    tableau(1, 12) = "I"
    tableau(1, 13) = "I"
    tableau(1, 14) = "O"
    tableau(1, 15) = "O"
    tableau(1, 16) = "O"
    tableau(1, 17) = "O"
    tableau(1, 18) = "O"
    tableau(1, 19) = "U"
    tableau(1, 20) = "U"
    tableau(1, 21) = "U"
    tableau(1, 22) = "U"
    tableau(1, 23) = "Y"
    tableau(1, 24) = "Y"
End Sub
