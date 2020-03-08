Attribute VB_Name = "lesMenus"
Option Explicit

'API functions
#If Win64 Then 'Pour utiliser sous Windows 64 bits
    Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As LongPtr
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As LongPtr
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else 'Pour utiliser sous Windows 32 bits
    Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Public Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

'Constantes
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const GWL_EXSTYLE = (-20)
Public Const HWND_TOP = 0
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_EX_APPWINDOW = &H40000
Public Const GWL_STYLE = (-16)
Public Const WS_MINIMIZEBOX = &H20000
Public Const SWP_FRAMECHANGED = &H20

Public param As UserForm

Public Sub CreerContextuel()
    Dim Barre As CommandBar
    Dim Controle As CommandBarControl
 
    Set Barre = CommandBars.Add(Name:="Contexte", Position:=msoBarPopup, Temporary:=True)
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Couper"
        .OnAction = "mnCouper"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Copier"
        .OnAction = "mnCopier"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Coller"
        .OnAction = "mnColler"
    End With
 
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Effacer la sélection"
        .OnAction = "mnSelectClear"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Effacer le contenu"
        .OnAction = "mnClearContent"
    End With
    
    Set Controle = Barre.Controls.Add
    With Controle
        .Caption = "Sélectionner tout"
        .OnAction = "mnSelectAll"
    End With
    
    Set Controle = Nothing
    Set Barre = Nothing
End Sub

Private Sub mnCouper()
    Dim donnee As New DataObject
    
    On Error Resume Next
    
    With param.ActiveControl
        donnee.SetText .SelText
        .SelText = ""
        donnee.PutInClipboard
    End With
End Sub

Private Sub mnCopier()
    Dim donnee As New DataObject

    On Error Resume Next
    
    donnee.SetText param.ActiveControl.SelText
    donnee.PutInClipboard
        
End Sub
 
Private Sub mnColler()
    Dim donnee As New DataObject

    On Error Resume Next
    
    donnee.GetFromClipboard
    param.ActiveControl.SelText = donnee.GetText
End Sub

Private Sub mnSelectClear()
    On Error Resume Next
    param.ActiveControl.SelText = ""
End Sub

Private Sub mnClearContent()
    On Error Resume Next
    param.ActiveControl.Value = ""
End Sub

Private Sub mnSelectAll()
    On Error Resume Next
    
    With param.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Value)
    End With
End Sub

'''''''''''''''''
'Ajout du bouton "réduire"
Public Sub AddMinimiseButton()
   Dim hwnd As Long
   hwnd = GetActiveWindow
   Call SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_MINIMIZEBOX)
   Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE)
End Sub

'Ajout d'un onglet dans la barre des tâches afin de faciliter l'accès
Public Sub AppTasklist(myForm)
   Dim WStyle As Long, Result As Long, hwnd As Long

   hwnd = FindWindow(vbNullString, myForm.Caption)
   WStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
   WStyle = WStyle Or WS_EX_APPWINDOW
   Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_HIDEWINDOW)
   Result = SetWindowLong(hwnd, GWL_EXSTYLE, WStyle)
   Result = SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
End Sub
