<div align="center">

## Generate a random string, with a length within a range\.


</div>

### Description

This code will take a range (lower and upper), and output a string of random characters (0-9, A-Z, a-z). I use this to generate a key for encryption, during the key negotiation phase of a connection to an encrypted server.

Usage is simple:

Dim sKey as string

sKey = GenerateKey(10,100)

' this code generates a key with a length between

' 10 and 100 characters.
 
### More Info
 
iLower: Lowest possible length for the string

iUpper: Maximum length for the string

Returns a string.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Gregg Housh](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gregg-housh.md)
**Level**          |Intermediate
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gregg-housh-generate-a-random-string-with-a-length-within-a-range__1-21395/archive/master.zip)





### Source Code

```
' Credit goes to these people for code I
' borrowed/modified:
' Kevin Lawrence - non-repeating random
' number generator
' VBPJ - GenerateRandomNumberInRange
' (modified by me, from a shuffle routine in VBPJ)
Public Function GenerateKey(ByVal iLower As Integer, ByVal iUpper As Integer) As String
  Dim sKey As String
  Dim sChar As String
  Dim iLen As Integer
  Dim iLoop As Integer
  ' dont need keys TOO big ...
  iLen = GetRandomNumberInRange(iLower, iUpper)
  For iLoop = 1 To iLen
    ' dont include quotes
Retry:
    Do
      sChar = Chr(GetRandomNumber())
    Loop While sChar = Chr(34)
    ' make sure its 0-9, A-Z, or a-z
    If Not IsValidChar(sChar) Then
      GoTo Retry:
    Else
      sKey = sKey & sChar
    End If
  Next iLoop
  GenerateKey = sKey
End Function
Private Function IsValidChar(ByVal sChar As String) As Boolean
  Dim btoggle As Boolean
  If Asc(sChar) >= 48 And Asc(sChar) <= 57 Then
    'valid #
    btoggle = True
  ElseIf Asc(sChar) >= 65 And Asc(sChar) <= 90 Then
    'valid uppercase character
    btoggle = True
  ElseIf Asc(sChar) >= 97 And Asc(sChar) <= 122 Then
    btoggle = True
  Else
    btoggle = False
  End If
  IsValidChar = btoggle
End Function
Public Function GetRandomNumberInRange(Lower As Integer, Upper As Integer) As Integer
  Static PrimeFactor(10) As Integer
  Static a As Integer
  Static c As Integer
  Static b As Integer
  Static s As Long
  Static n As Integer
  Static n1 As Integer
  Dim i As Integer
  Dim j As Integer
  Dim K As Integer
  Dim m As Integer
  Dim t As Boolean
  If (n <> Upper - Lower + 1) Then
    n = Upper - Lower + 1
    i = 0
    n1 = n
    K = 2
    Do While K <= n1
      If (n1 Mod K = 0) Then
        If (i = 0 Or PrimeFactor(i) <> K) Then
          i = i + 1
          PrimeFactor(i) = K
        End If
        n1 = n1 / K
      Else
        K = K + 1
      End If
    Loop
    b = 1
    For j = 1 To i
      b = b * PrimeFactor(j)
    Next j
    If n Mod 4 = 0 Then b = b * 2
    a = b + 1
    c = Int(n * 0.66)
    t = True
    Do While t
      t = False
      For j = 1 To i
        If ((c Mod PrimeFactor(j) = 0) Or (c Mod a = 0)) Then t = True
      Next j
      If t Then c = c - 1
    Loop
    Randomize
    s = Rnd(n)
  End If
  s = (a * s + c) Mod n
  GetRandomNumberInRange = s + Lower
End Function
Public Function GetRandomNumber() As Integer
    Dim a(122) ' Sets the maximum number To pick
    Dim b(122) ' Will be the list of new numbers (same as DIM above)
    Dim ChosenNumber As Integer
    Dim MaxNumber As Integer
    Dim seq As Integer
    'Set the original array
    MaxNumber = 122 ' Must equal the Dim above
    For seq = 0 To MaxNumber
      a(seq) = seq
    Next seq
    'Main Loop (mix em all up)
    Randomize (Timer)
    For seq = MaxNumber To 0 Step -1
      ChosenNumber = Int(seq * Rnd)
      b(MaxNumber - seq) = a(ChosenNumber)
      a(ChosenNumber) = a(seq)
    Next seq
  ' return a random number from a random position in B()
  GetRandomNumber = b(GetRandomNumberInRange(1, 122))
End Function
```

