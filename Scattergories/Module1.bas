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
    
    tableau(0, 0) = "�"
    tableau(0, 1) = "�"
    tableau(0, 2) = "�"
    tableau(0, 3) = "�"
    tableau(0, 4) = "�"
    tableau(0, 5) = "�"
    tableau(0, 6) = "�"
    tableau(0, 7) = "�"
    tableau(0, 8) = "�"
    tableau(0, 9) = "�"
    tableau(0, 10) = "�"
    tableau(0, 11) = "�"
    tableau(0, 12) = "�"
    tableau(0, 13) = "�"
    tableau(0, 14) = "�"
    tableau(0, 15) = "�"
    tableau(0, 16) = "�"
    tableau(0, 17) = "�"
    tableau(0, 18) = "�"
    tableau(0, 19) = "�"
    tableau(0, 20) = "�"
    tableau(0, 21) = "�"
    tableau(0, 22) = "�"
    tableau(0, 23) = "�"
    tableau(0, 24) = "�"
    
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
