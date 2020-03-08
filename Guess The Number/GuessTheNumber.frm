VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GuessTheNumber 
   Caption         =   "Guess the number - Hugues DUMONT"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "GuessTheNumber.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GuessTheNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.answer.Visible = False
    Me.newGuess.Visible = False
    Me.guess.value = ""
    Me.validate.Enabled = True
    Me.answer.Caption = CStr(intRndBetween(0, 10000))
    Call getTips(CLng(Me.answer.Caption))
End Sub

Private Sub newGuess_Click()
 Call UserForm_Initialize
End Sub

Private Sub validate_Click()
    If (Not isInteger(Me.guess.value)) Then
        MsgBox "You need guess a valid integer before validating.", vbOKOnly + vbExclamation, "Missing or incorrect value"
        Exit Sub
    End If
    
    If (Me.answer.Caption = CStr(Me.guess.value)) Then
        Me.answer.Caption = "You guessed right : " & Me.answer.Caption
        Me.answer.ForeColor = RGB(0, 255, 0)
    Else
        Me.answer.Caption = "You are wrong : " & Me.answer.Caption
        Me.answer.ForeColor = RGB(255, 0, 0)
    End If
    Me.validate.Enabled = False
    Me.answer.Visible = True
    Me.newGuess.Visible = True
End Sub

'Function Calling all the other functions to get informations about the number to guess
Private Sub getTips(number As Long)
    Dim prime As Boolean
    Dim primeFac() As String
    
    prime = isPrime(number)
    Me.tips.Caption = IIf(prime, "- Is a prime number", "- Is not prime")
    Me.tips.Caption = Me.tips.Caption & Chr(13) & IIf(isPerfectNumber(number), "- Is a perfect number", "- Is not perfect")
    Me.tips.Caption = Me.tips.Caption & Chr(13) & "- Add " & CStr(nextPrime(number) - number) & " to get next prime"
    
    If (number > 2) Then
        Me.tips.Caption = Me.tips.Caption & Chr(13) & "- Substract " & CStr(number - previousPrime(number)) & " to get previous prime"
    End If
    
    If (Not prime) Then
        primeFac = Split(primeFactors(number), ",")
        Me.tips.Caption = Me.tips.Caption & Chr(13) & "- Biggest prime factor is " & Replace(primeFac(UBound(primeFac)), " ", "")
        If (number > 1) Then
            Me.tips.Caption = Me.tips.Caption & Chr(13) & "- It has " & UBound(Split(factors(number), " ")) + 1 & " factors, including 1 and itself"
            Me.tips.Caption = Me.tips.Caption & Chr(13) & "- It has " & UBound(Split(primeFactors(number), " ")) + 1 & " prime factors"
        End If
    End If
    
    Me.tips.Caption = Me.tips.Caption & Chr(13) & "- If you add all its digits once you get the number " & CStr(sumDigitsOnce(number))
    Me.tips.Caption = Me.tips.Caption & Chr(13) & "- If you add all its digits until only one is left you get the digit " & CStr(sumAllDigits(number))
    Me.tips.Caption = Me.tips.Caption & Chr(13) & IIf(Int(Sqr(number)) ^ Int(Sqr(number)) = number, "- It's square root is an integer", "- It's square root is a decimal number")
    
    If (Len(CStr(number)) > 2) Then
        Me.tips.Caption = Me.tips.Caption & Chr(13) & "- First digit is " & Mid(CStr(number), 1, 1)
    End If
End Sub

'Function to generate an integer (int) between 2 values (min, max)
Private Function intRndBetween(min As Integer, max As Integer) As Integer
    Randomize
    intRndBetween = Int((max - min + 1) * Rnd + min)
End Function

'Function to check if string is an integer
Private Function isInteger(value As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    Const INT_MIN As Integer = -32768
    Const INT_MAX As Integer = 32767
    
    reg.Pattern = "^(-)?(\d)+$"
    isInteger = False
    If reg.test(value) Then
        On Error GoTo capacityOverflow
        If CInt(value) >= INT_MIN And CInt(value) <= INT_MAX Then
            isInteger = True
        End If
    End If
    Set reg = Nothing
    Exit Function
capacityOverflow:
    MsgBox "Value is integer but over 32 767 or lower than -32 768" & Chr(13) & _
        "Can't be converted to the integer type in vba (might be able to convert to long type)!", _
            vbOKOnly + vbCritical, "Capacity overflow !"
    Set reg = Nothing
End Function
