Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public start As Single

Public Function lePlusLong(tirage As String) As Long
    Dim ws As Worksheet
    Dim used(10) As String
    Dim i As Long, k As Long, col As Long, lig As Long, maxLen As Long
    Dim trouve As Boolean
    
    Set ws = ThisWorkbook.Sheets("MOTS")
    
    maxLen = 0
    trouve = False
    For col = 10 To 1 Step -1
        If maxLen > 0 Then Exit For
        For lig = 1 To ws.Cells(Rows.Count, col).End(xlUp).Row
            For i = 1 To 10
                used(i) = Mid(tirage, i, 1)
            Next i
    
            For i = 1 To Len(ws.Cells(lig, col).Value)
                trouve = False
                For k = 1 To 10
                    If used(k) = Mid(ws.Cells(lig, col).Value, i, 1) Then
                        used(k) = ""
                        trouve = True
                        Exit For
                    End If
                Next k
                If Not trouve Then Exit For
            Next i
            If trouve Then
                UserForm1.ListBox1.AddItem ws.Cells(lig, col).Value
                maxLen = col
            End If
            UserForm1.Label1.Caption = (30 - Round(Timer - start, 0)) & " s"
            DoEvents
        Next lig
        UserForm1.Label1.Caption = (30 - Round(Timer - start, 0)) & " s"
        DoEvents
    Next col
    lePlusLong = maxLen
End Function
