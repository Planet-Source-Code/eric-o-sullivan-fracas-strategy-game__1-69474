Attribute VB_Name = "NameGen"
Option Explicit
Option Base 1

Public NameSpice(25) As String
Dim LetterPrp(26) As Integer
Dim LetterPct(26) As Single

Public Function GenerateName(MinSize As Integer, MaxSize As Integer, Specials As Boolean)

Dim WorkingName As String
Dim LastChar As String
Dim TwoCharsAgo As String
Dim GoodWord As Boolean
Dim i As Long

Do

'The first character has an equal chance of being any letter.
WorkingName = Chr(Int(Rnd(1) * 26) + 65)
LastChar = WorkingName
TwoCharsAgo = ""
GoodWord = True   'For now.

For i = 1 To (Int(Rnd(1) * (MaxSize - MinSize + 1)) + MinSize) - 1


  'Let's see if we had a consonant or a vowel last.
  If Consonant(LastChar) = True Then
    'We had a consonant last time.
    If Consonant(TwoCharsAgo) = True Then
      'We have two consonants in a row.  This one must be a vowel,
      'Unless the last one was a Y.
      If LastChar = "Y" Then
        WorkingName = WorkingName & AddChar("BCDFGKLMNPRSTVXZ")
      Else
        WorkingName = WorkingName & AddChar("AEIOU")
      End If
    Else
      'We only have one consonant so far, so there's a chance that
      'the next character could be one too.
      If Rnd(1) < 0.2 Then
        'It's a vowel this time.
        'Let's make sure that a U follows a Q...
        If LastChar = "Q" Then
          WorkingName = WorkingName & AddChar("U")
        Else
          WorkingName = WorkingName & AddChar("AEIOU")
        End If
      Else
        'Let's add a sensible consonant after this one.
        Select Case LastChar
        Case "B":
          WorkingName = WorkingName & AddChar("BLRSY")
        Case "C":
          WorkingName = WorkingName & AddChar("CHKLRSTY")
        Case "D":
          WorkingName = WorkingName & AddChar("DRSWY")
        Case "F":
          WorkingName = WorkingName & AddChar("FLRSTY")
        Case "G":
          WorkingName = WorkingName & AddChar("GHLRS")
        Case "H":
          WorkingName = WorkingName & AddChar("AEIOUY")
        Case "J":
          WorkingName = WorkingName & AddChar("AEIOU")
        Case "K":
          WorkingName = WorkingName & AddChar("HKLNRSTW")
        Case "L":
          WorkingName = WorkingName & AddChar("CDKLMNPSTY")
        Case "M":
          WorkingName = WorkingName & AddChar("MPSY")
        Case "N":
          WorkingName = WorkingName & AddChar("CDKNSTY")
        Case "P":
          WorkingName = WorkingName & AddChar("HLPRSTY")
        Case "Q":
          WorkingName = WorkingName & AddChar("U")
        Case "R":
          WorkingName = WorkingName & AddChar("CDKMNPRSTY")
        Case "S":
          WorkingName = WorkingName & AddChar("CHKLMNPSTWY")
        Case "T":
          WorkingName = WorkingName & AddChar("HRSTWY")
        Case "V":
          WorkingName = WorkingName & AddChar("LRS")
        Case "W":
          WorkingName = WorkingName & AddChar("HKST")
        Case "X":
          WorkingName = WorkingName & AddChar("CSYAEIOU")
        Case "Y":
          WorkingName = WorkingName & AddChar("LSTKMNPC")
        Case "Z":
          WorkingName = WorkingName & AddChar("HZZZ")
        End Select
      End If
    End If
  Else
    'We had a vowel last time.
    If Consonant(TwoCharsAgo) = True Then
      'We only have one vowel so far, so there's a chance that
      'the next character could be a vowel too.
      If (Rnd(1) < 0.9) And Right(WorkingName, 2) <> "QU" Then
        'It's a consonant this time.
        WorkingName = WorkingName & AddChar("BCDFGHJKLMNPQRSTVWX")
      Else
        'Let's add a sensible vowel after this one.
        Select Case LastChar
        Case "A":
          WorkingName = WorkingName & AddChar("IU")
        Case "E":
          WorkingName = WorkingName & AddChar("AEIOU")
        Case "I":
          WorkingName = WorkingName & AddChar("AEO")
        Case "O":
          WorkingName = WorkingName & AddChar("AEIOU")
        Case "U":
          WorkingName = WorkingName & AddChar("AEI")
        End Select
      End If
    Else
      'We have two vowels in a row.  This one must be a consonant.
      'Note that some consonants don't work well after two vowels.
      WorkingName = WorkingName & AddChar("BCDFGKLMNPRST")
    End If
  End If

TwoCharsAgo = LastChar
LastChar = Right(WorkingName, 1)

Next i

'We need to double-check some illegal letter combinations at the start.
TwoCharsAgo = Left(WorkingName, 1)    'First char.
LastChar = Mid(WorkingName, 2, 1)     'Second char.
'If there's a double-consonant at the beginning, that's bad.
If TwoCharsAgo = LastChar Then GoodWord = False
'If there's a consonant followed by S or D, that's bad.
If LastChar = "S" Or LastChar = "D" Then
  If Consonant(TwoCharsAgo) = True Then GoodWord = False
End If
'If there's a consonant followed by T, K, M, N, P, or C, that's bad.
'Unless an S starts the word, of course.
If LastChar = "T" Or LastChar = "K" Or LastChar = "M" Or LastChar = "N" _
                  Or LastChar = "P" Or LastChar = "C" Then
   If Consonant(TwoCharsAgo) = True And TwoCharsAgo <> "S" Then GoodWord = False
End If
'If this word starts with a Y, it had better not have a consonant after it.
If TwoCharsAgo = "Y" And Consonant(LastChar) = True Then GoodWord = False
'No QU at the end.  Yuck.
If Right(WorkingName, 2) = "QU" Then GoodWord = False
'IY, YI and IW sure look stupid.
If InStr(WorkingName, "IY") > 0 Or InStr(WorkingName, "IW") > 0 Or InStr(WorkingName, "YI") > 0 Then GoodWord = False
'So do UY and UW.
If InStr(WorkingName, "UY") > 0 Or InStr(WorkingName, "UW") > 0 Then GoodWord = False
'And while we're at it, so do IH and UH.
If InStr(WorkingName, "IH") > 0 Or InStr(WorkingName, "UH") > 0 Then GoodWord = False
'Words that start with X or IL generally suck, I find.
If TwoCharsAgo = "X" Or Left(WorkingName, 2) = "IL" Then GoodWord = False

Loop Until GoodWord = True

LastChar = Right(WorkingName, 1)
TwoCharsAgo = Mid(WorkingName, Len(WorkingName) - 1, 1)   '2nd to last char.

'Now we will spice up the name with some embellishments.
If ((Consonant(LastChar) = True) And (Consonant(TwoCharsAgo) = True)) Or _
   (Len(WorkingName) < 4) Or (Rnd(1) < 0.15) Then
  'Let's add something cool to the end of the name!
  If (Consonant(LastChar) = True) And (UCase(LastChar) <> "Y") Then
    WorkingName = WorkingName & NameSpice(Int(Rnd(1) * 6) + 17)
  Else
    WorkingName = WorkingName & NameSpice(Int(Rnd(1) * 6) + 11)
  End If
End If

'Let's put the proper case on this word.
WorkingName = UCase(Left(WorkingName, 1)) & LCase(Right(WorkingName, Len(WorkingName) - 1))

If Specials = True Then

  If Rnd(1) < 0.13 Then
    'Let's give this country a formal title!
    WorkingName = NameSpice(Int(Rnd(1) * 10) + 1) & " " & WorkingName
  End If

End If

'That's it!  We've got one cool name for you.
GenerateName = WorkingName

End Function

Public Function Consonant(CharIn As String) As Boolean

If CharIn = "A" Or CharIn = "E" Or CharIn = "I" Or CharIn = "O" Or CharIn = "U" Or (Len(CharIn) <> 1) Then
  Consonant = False
Else
  Consonant = True
End If

End Function

Public Function AddChar(Letters As String)

Dim j As Long
Dim LetterTotal As Single
Dim LetterThresh As Single
Dim RandLetter As Single

'Let's figure out our proportions.
LetterTotal = 0
For j = 1 To Len(Letters)
  LetterTotal = LetterTotal + LetterPrp(Asc(Mid(Letters, j, 1)) - 64)
Next j

'Let's set up our threshholds for each character in the Letters string.
LetterThresh = 0
For j = 1 To Len(Letters)
  LetterThresh = LetterThresh + (LetterPrp(Asc(Mid(Letters, j, 1)) - 64) / LetterTotal)
  LetterPct(j) = LetterThresh
Next j

'Now grab one of them.
RandLetter = Rnd(1)
For j = 1 To Len(Letters)
  If RandLetter <= LetterPct(j) Then
    'It's this one!
    AddChar = Mid(Letters, j, 1)
    Exit For
  End If
Next j

End Function

Public Sub InitNameStuff()

'We will use the Scrabble(TM) letter proportions!
LetterPrp(1) = 9
LetterPrp(2) = 2
LetterPrp(3) = 2
LetterPrp(4) = 4
LetterPrp(5) = 12
LetterPrp(6) = 2
LetterPrp(7) = 3
LetterPrp(8) = 2
LetterPrp(9) = 9
LetterPrp(10) = 1
LetterPrp(11) = 1
LetterPrp(12) = 4
LetterPrp(13) = 2
LetterPrp(14) = 6
LetterPrp(15) = 8
LetterPrp(16) = 2
LetterPrp(17) = 1
LetterPrp(18) = 6
LetterPrp(19) = 4
LetterPrp(20) = 6
LetterPrp(21) = 4
LetterPrp(22) = 2
LetterPrp(23) = 2
LetterPrp(24) = 1
LetterPrp(25) = 2
LetterPrp(26) = 1

NameSpice(1) = "Upper"
NameSpice(2) = "Lower"
NameSpice(3) = "Old"
NameSpice(4) = "New"
NameSpice(5) = "San"
NameSpice(6) = "Costa"
NameSpice(7) = "North"
NameSpice(8) = "South"
NameSpice(9) = "East"
NameSpice(10) = "West"
NameSpice(11) = "tia"
NameSpice(12) = "lia"
NameSpice(13) = "way"
NameSpice(14) = "land"
NameSpice(15) = "ton"
NameSpice(16) = "burg"
NameSpice(17) = "ary"
NameSpice(18) = "age"
NameSpice(19) = "ia"
NameSpice(20) = "any"
NameSpice(21) = "ica"
NameSpice(22) = "ania"
NameSpice(23) = "Isle"
NameSpice(24) = "Island"
NameSpice(25) = "The Isle of"
End Sub

