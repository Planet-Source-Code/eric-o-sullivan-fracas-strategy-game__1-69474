Attribute VB_Name = "GraphicFX"
Option Explicit
Option Base 1

Public Sub DrawMap()

'These commands draw the map.  MapBuffer is where the map is drawn.
MyMap.DisplayMap Land!LandMap.hdc, Land!MapBuffer.hdc

End Sub

Public Sub FlashTimerSub()

'This sub cycles through the flash array and updates everything.

If (GameMode <> GM_GAME_ACTIVE) Or (CFGFlashing = 0) Then Exit Sub

Dim i As Long
Dim FlagUpdate As Boolean

FlagUpdate = False
For i = 1 To MyMap.NumberOfCountries
  'Moved this into the map class so there's only one access.
  If (MyMap.UpdateFlashing(i)) = True Then FlagUpdate = True
Next i

If FlagUpdate = True Then DrawMap

End Sub

Public Sub ResetFlashing()

Dim i As Long
Dim j As Integer

'This sub clears out the flash array.

For i = 1 To MyMap.NumberOfCountries
  For j = 1 To 4
    MyMap.Flash(i, j) = 0
  Next j
Next i

End Sub

Public Sub ShortFlash(Cntry As Long)

'This sub starts off a short single flash for the passed country.

If CFGFlashing = 0 Then Exit Sub
MyMap.Flash(Cntry, 1) = 1
MyMap.Flash(Cntry, 2) = 6
MyMap.Flash(Cntry, 3) = 6
MyMap.Flash(Cntry, 4) = 1

End Sub

Public Sub LongFlash(Cntry As Long)

'This sub starts off a long single flash for the passed country.

If CFGFlashing = 0 Then Exit Sub
MyMap.Flash(Cntry, 1) = 1
MyMap.Flash(Cntry, 2) = 12
MyMap.Flash(Cntry, 3) = 12
MyMap.Flash(Cntry, 4) = 1

End Sub

Public Sub LongerFlash(Cntry As Long)

'This sub starts off a long single flash for the passed country.

If CFGFlashing = 0 Then Exit Sub
MyMap.Flash(Cntry, 1) = 1
MyMap.Flash(Cntry, 2) = 30
MyMap.Flash(Cntry, 3) = 30
MyMap.Flash(Cntry, 4) = 1

End Sub
