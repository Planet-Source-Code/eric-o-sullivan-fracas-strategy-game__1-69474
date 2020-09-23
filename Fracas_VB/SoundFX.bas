Attribute VB_Name = "SoundFX"
Option Explicit
Option Base 1

Public Function PlayWav(strPath As String, sndVal As Long)

'Don't play a sound if sfx are disabled.
If CFGSound = 2 Then Exit Function

sndPlaySound SFXpath & strPath, sndVal

End Function

Public Sub SNDTroopsIn(PlayrNum As Integer)
  PlayWav "dn" & Trim(Str(PlayrNum)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDTroopsOut(PlayrNum As Integer)
  PlayWav "up" & Trim(Str(PlayrNum)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDBuildAPort()
  PlayWav "port" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDPlayWarCadence()
  If Rnd(1) < 0.3 Then
    PlayWav "timpani1.wav", SND_GAME_FLAGS
  Else
    PlayWav "timpani2.wav", SND_GAME_FLAGS
  End If
End Sub

Public Sub SNDPlayFanfare(FanfareNum As Integer)
  PlayWav "fanfare" & Trim(Str(FanfareNum)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDSmallExplosion()
  PlayWav "smexp" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDMediumExplosion()
  PlayWav "mdexp" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDLargeExplosion()
  PlayWav "lgexp" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDWhistle()
  PlayWav "ladeda" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDApplause()
  PlayWav "yay" & Trim(Str(Int(Rnd(1) * 2) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDSplishSplash()
  PlayWav "splash" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub

Public Sub SNDBooBoo()
  PlayWav "booboo.wav", SND_GAME_FLAGS
End Sub

Public Sub SNDBonusTwinkles()
  PlayWav "bonus" & Trim(Str(Int(Rnd(1) * 3) + 1)) & ".wav", SND_GAME_FLAGS
End Sub
