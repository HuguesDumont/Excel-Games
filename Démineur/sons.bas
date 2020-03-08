Attribute VB_Name = "sons"
Option Explicit

Private Const SND_FILENAME = &H20000
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Sub jouerSon(fichier As String)
    PlaySound ThisWorkbook.Path & "\" & fichier & ".wav", ByVal SND_SYNC, SND_FILENAME Or SND_ASYNC
End Sub

