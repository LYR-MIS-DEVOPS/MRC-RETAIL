VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKeyTrapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum KeyTrapping
    NumbersOnly
    NumbersWithDecimalPoint
    ValidChars
    ValidSQLChars
    ValidCodeChars
    ValidCodeCharsWithSpace
    ValidPassword
End Enum

Public Sub Validation(ByRef objtxt As Object, ByRef KeyAscii As Integer, Direction As KeyTrapping)
    Dim str As String
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then Exit Sub
 
    If Direction = KeyTrapping.NumbersOnly Then
    
        If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If
        
    ElseIf Direction = KeyTrapping.NumbersWithDecimalPoint Then
    
        If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = Asc(".") Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If
        
    ElseIf Direction = KeyTrapping.ValidChars Then
    
        If (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = Asc(".") Or KeyAscii = Asc("_") Or KeyAscii = Asc("-") _
            Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc(" ") Or KeyAscii = Asc("#") Or KeyAscii = Asc(",") _
            Or KeyAscii = Asc("&") Or KeyAscii = Asc("!") Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If
        
    ElseIf Direction = KeyTrapping.ValidSQLChars Then
    
        If KeyAscii = Asc("'") Or KeyAscii = Asc("""") Then
            KeyAscii = 0
            Beep!
        Else
            Exit Sub
        End If
        
        
    ElseIf Direction = ValidCodeChars Then
    
        If (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = Asc(".") Or KeyAscii = Asc("_") Or KeyAscii = Asc("-") _
            Or KeyAscii = Asc("#") Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If

    ElseIf Direction = ValidCodeCharsWithSpace Then
    
        If (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = Asc(".") Or KeyAscii = Asc("_") Or KeyAscii = Asc("-") _
            Or KeyAscii = Asc("#") Or KeyAscii = Asc(" ") Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If

    ElseIf Direction = ValidPassword Then
    
        If (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) _
            Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = Asc(".") Or KeyAscii = Asc("_") Or KeyAscii = Asc("-") _
            Or KeyAscii = Asc("~") Or KeyAscii = Asc("`") Or KeyAscii = Asc("!") Or KeyAscii = Asc("@") Or KeyAscii = Asc("$") Or _
            KeyAscii = Asc("%") Or KeyAscii = Asc("^") Or KeyAscii = Asc("&") Or KeyAscii = Asc("*") Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") _
            Or KeyAscii = Asc("=") Or KeyAscii = Asc("+") Then
            Exit Sub
        Else
            KeyAscii = 0
            Beep!
        End If

    Else
        KeyAscii = 0
        Beep!
    End If
End Sub

