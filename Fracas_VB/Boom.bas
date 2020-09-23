Attribute VB_Name = "Boom"
'This form contains code from Boom! Particle Explosion Simulation.
'Written by Jason Merlo 10/26/99.

Option Explicit
Option Base 1

'The balls array contains the vital statistics of each ball.
Dim Balls(BALLMAX) As Ball

'Water variables dimensions.
Public WaterColors(6) As Long
Dim Waves(WAVEMAX) As Wave

Private Sub PrepareBallBuffer()

'This sub simply copies the MapBuffer background picture into the PicBuffer.

Land!PicBuffer.Cls   'Get rid of that stinkin' edging problem!
u% = BitBlt(Land!PicBuffer.hdc, 0, 0, Wide, Tall, Land!MapBuffer.hdc, 0, 0, SRCCOPY)

End Sub

Private Function DrawBall(i As Integer, Source As Long, Dest As Long, Shape As Integer, Color As Integer, x As Long, y As Long)

'Now we draw the new ball.
If Balls(i).Shape < 5 Then
  u% = BitBlt(Dest, x, y, 5, 5, Source, (Shape - 1) * 7, 91, SRCAND)
  u% = BitBlt(Dest, x, y, 5, 5, Source, (Shape - 1) * 7, (Color * 7), SRCINVERT)
Else
  u% = BitBlt(Dest, x - 2, y, 7, 7, Source, (Shape - 1) * 7, 91, SRCAND)
  u% = BitBlt(Dest, x - 2, y, 7, 7, Source, (Shape - 1) * 7, (Color * 7), SRCINVERT)
End If

End Function

Private Function DrawShadow(i As Integer, Source As Long, Dest As Long, Shape As Integer, x As Long, Newshad As Long)

'Draw the shadow so that the ball will overlap it.
If Balls(i).Shape < 5 Then
  u% = BitBlt(Dest, x, Newshad, 5, 3, Source, (Shape - 1) * 7, 100, SRCAND)
  u% = BitBlt(Dest, x, Newshad, 5, 3, Source, (Shape - 1) * 7, 97, SRCINVERT)
Else
  u% = BitBlt(Dest, x - 2, Newshad + 5, 7, 3, Source, (Shape - 1) * 7, 101, SRCAND)
  u% = BitBlt(Dest, x - 2, Newshad + 5, 7, 3, Source, (Shape - 1) * 7, 98, SRCINVERT)
End If

End Function

Public Sub BallTimerSub()

Dim i As Integer
Dim Absorb As Single
Dim Xpos As Single
Dim Ypos As Single
Dim Yvel As Single
Dim Ytilt As Single
Dim Ytiltvel As Single
Dim Ystartloc As Single
Dim Yshadow As Long

'This is where we modify each ball's position based on
'the explosion's properties and global settings.

'Step 1:  Copy the blank background over.
Call PrepareBallBuffer

'Step 2:  Ball physics.
For i = 1 To BALLMAX
  'Get out of this loop if there are no balls -- the most we'll have
  'is a single selection arrow.
  If (i > 2) And (CFGExplosions = 0) Then Exit For
  'Only operate on balls that are still bouncing.
  If Balls(i).Enabled = True Then
    'Get the class properties so we can work locally.  Now the values
    'in the class are the 'old' values, used for erasing balls
    'and for shadow calculations.
    Xpos = Balls(i).Xpos
    Ypos = Balls(i).Ypos
    Yvel = Balls(i).Yvel
    Ytilt = Balls(i).Ytilt
    Ytiltvel = Balls(i).Ytiltvel
    Ystartloc = Balls(i).Ystart
    Absorb = Balls(i).Elastic
    'Apply gravity to the vertical velocity.
    Yvel = Yvel + GRAVITY
    'Adjust the tilt of the ball. This is used to skew the pattern for
    'a 3-d type of effect.  A z-axis modifier, if you will.
    Ytilt = Ytilt + Ytiltvel
    'Now check if we can kill this ball because it's out of bounds.
    'Note that the -5 on the Xpos check is because that is the width
    'of a single ball!  Just a fudge since they were hanging over to the right.
    If (Xpos < 0) Or (Xpos > Wide - 5) Or (Balls(i).Yshadow < 0) Or ((Ypos + Ytilt + Ytiltvel) > Tall) Then
      Balls(i).Enabled = False
    Else
      'Move ball.
      Xpos = Xpos + Balls(i).Xvel
      Ypos = Ypos + Yvel
      'Calculate the shadow position.
      Yshadow = (Int(Balls(i).Ypos + Ytilt) + Int(Ystartloc - Balls(i).Ypos)) + 2
      'If we went past our starting point, we need to rebound.
      If Ypos > Ystartloc Then
        'The ground absorbs some velocity and reverses the ball's direction.
        Yvel = Absorb * (-Yvel)
        Ypos = Ystartloc
        'If the ball has slowed down enough, or if it has hit water,
        'we will stop it altogether.
        If Abs(Yvel) < (0.5 * GRAVITY) Then
          'Take this ball out of service and free up a slot for a new one.
          Balls(i).Enabled = False
        End If
        If (MyMap.Grid(((Xpos - 2) / 8) + 1, ((Yshadow - 2) / 8) + 1) > TILEVAL_COASTLINE) Then
          'Take this ball out of service and free up a slot for a new one...
          Balls(i).Enabled = False
          '...and make a splash since we hit water!
          Call MakeSplash(Int(Xpos), Int(Yshadow))
        End If
      End If
    End If
    'If this ball didn't die, we need to get its shadow position
    'and copy its stats over to the class properties.
    If Balls(i).Enabled = True Then
      'Update the values of all the parameters in the class.  These will
      'be the 'old' values on the next scan!
      Balls(i).Yshadow = Yshadow
      Balls(i).Xpos = Xpos
      Balls(i).Ypos = Ypos
      Balls(i).Yvel = Yvel
      Balls(i).Ytilt = Ytilt
      Balls(i).Ytiltvel = Ytiltvel
      Balls(i).Ystart = Ystartloc
    End If
  End If
Next i

'Now update the video buffer.  Note that balls are only drawn to PicBuffer.

'Step 3:  Draw all the shadows FIRST so that they don't overlap any balls.
For i = 1 To BALLMAX
  'Leave if we're only doing selection arrows.
  If (i > 2) And (CFGExplosions = 0) Then Exit For
  If Balls(i).Enabled = True Then
    If (Balls(i).Shape < 5) Or ((Balls(i).Shape = 5) And (Balls(i).Yvel > 0)) Or _
                               ((Balls(i).Shape = 6) And (Balls(i).Yvel < 0)) Or _
                               (Balls(i).Shape > 6) Then
      Call DrawShadow(i, Land!BallPic.hdc, Land!PicBuffer.hdc, Balls(i).Shape, Int(Balls(i).Xpos), Balls(i).Yshadow)
    End If
  End If
Next i

'Step 4:  Draw all the balls.
For i = 1 To BALLMAX
  'Leave if we're only doing selection arrows.
  If (i > 2) And (CFGExplosions = 0) Then Exit For
  If Balls(i).Enabled = True Then
    If (Balls(i).Shape < 5) Or ((Balls(i).Shape = 5) And (Balls(i).Yvel > 0)) Or _
                               ((Balls(i).Shape = 6) And (Balls(i).Yvel < 0)) Or _
                               (Balls(i).Shape > 6) Then
      Call DrawBall(i, Land!BallPic.hdc, Land!PicBuffer.hdc, Balls(i).Shape, Balls(i).Color, Int(Balls(i).Xpos), Int(Balls(i).Ypos + Balls(i).Ytilt))
    End If
  End If
Next i

'Leave if we don't want waves.
If CFGWaves > 0 Then
  'Animate the water.
  AnimateWater
End If

'Draw the screen once a scan to keep it fresh and clean.
Call Land.Form_Paint

End Sub

Public Sub AnimateWater()

Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim Resp As Long

For i = 1 To WAVEMAX
  'Let's see if we can start a random new wave.
  If Waves(i).Enabled = False And Rnd(1) < 0.02 Then
    x = Int(Rnd(1) * MyMap.Xsize) + 1
    y = Int(Rnd(1) * MyMap.Ysize) + 1
    
    If MyMap.Grid(x, y) > TILEVAL_COASTLINE Then
      'Found a slot with water!  Make a wave here.
      Waves(i).Enabled = True
      Waves(i).Frame = 0
      Waves(i).Count = 0
      Waves(i).Speed = Rnd(1) * 3 + 1
      Waves(i).Shape = Int(Rnd(1) * 8)
      Waves(i).Xpos = ((x - 1) * 8) + Int(Rnd(1) * 3)
      Waves(i).Ypos = ((y - 1) * 8) + Int(Rnd(1) * 7)
    End If
    
  ElseIf Waves(i).Enabled = True Then
    'Animate this one!
    Waves(i).Count = Waves(i).Count + 1
    If Waves(i).Count >= Waves(i).Speed Then
      'We advance the frame.
      Waves(i).Count = Waves(i).Count - Waves(i).Speed
      Waves(i).Frame = Waves(i).Frame + 1
    End If
    'Now draw the wave.
    If Waves(i).Frame < 6 Then
      u% = BitBlt(Land!PicBuffer.hdc, Waves(i).Xpos, Waves(i).Ypos, 5, 1, Land!WaterPic.hdc, (Waves(i).Frame - 1) * 5, (Waves(i).Shape * 2) + 1, SRCAND)
      u% = BitBlt(Land!PicBuffer.hdc, Waves(i).Xpos, Waves(i).Ypos, 5, 1, Land!WaterPic.hdc, (Waves(i).Frame - 1) * 5, Waves(i).Shape * 2, SRCINVERT)
    Else
      'This wave is done.
      Waves(i).Enabled = False
    End If
  End If
Next i

End Sub

Public Sub MakeSplash(xxx As Long, yyy As Long)

'This function is called whenever a ball hits water.

Dim j As Integer

If CFGWaves = 0 Then Exit Sub

For j = 1 To WAVEMAX
  If Waves(j).Enabled = False Then
    'We'll put the splash at j.
    Waves(j).Enabled = True
    Waves(j).Frame = 0
    Waves(j).Count = 0
    Waves(j).Speed = Rnd(1) * 5 + 1
    Waves(j).Shape = Int(Rnd(1) * 2) + 8
    Waves(j).Xpos = xxx
    Waves(j).Ypos = yyy
    Exit For
  End If
Next j

End Sub

Public Sub ClearAllBalls()

Dim i As Integer

'This sub cleans out the whole ball array.
For i = 1 To BALLMAX
  Set Balls(i) = New Ball
  Balls(i).Enabled = False
Next i

End Sub

Public Sub ClearAllWaves()

Dim i As Integer

'Clean out the wave array.
For i = 1 To WAVEMAX
  Set Waves(i) = New Wave
  Waves(i).Enabled = False
Next i

End Sub

Public Sub BuildBalls(StartNum As Long, Xstart As Long, Ystart As Long, Intensity As Long, Spread As Long, AbsorbPct As Long, Size As Integer, Color As Integer)

'This sub creates an explosion!

Dim i As Integer
Dim j As Integer

'Leave unless we know we're making a selection arrow.
If (CFGExplosions = 0) And ((Size < 5) Or (Size > 6)) Then Exit Sub

'Do this for each new ball.
For j = 1 To StartNum

  'Find an empty ball slot and fill it up with info.
  For i = 1 To BALLMAX
    If Not Balls(i).Enabled Then
      'Slot i is free.  Let's populate it.
      Set Balls(i) = New Ball
      Balls(i).Enabled = True
      Balls(i).Xpos = Xstart
      Balls(i).Ypos = Ystart
      Balls(i).Xvel = (Rnd(1) * (Spread / 5)) - ((Spread / 5) / 2)
      If Intensity = -1 Then 'Used for Computer turns.
        Balls(i).Yvel = -11.37
      Else
        Balls(i).Yvel = -((Rnd(1) * Intensity / 3) + (Intensity / 20))
      End If
      Balls(i).Ystart = Ystart
      Balls(i).Ytilt = 0
      Balls(i).Ytiltvel = (Rnd(1) * (Spread / 5)) - ((Spread / 5) / 2)
      Balls(i).Yshadow = Ystart + 2
      Select Case Size
        Case 5
            'Used for Computer turns.
            Balls(i).Shape = 5
        Case 6
            'Used for Computer turns.
            Balls(i).Shape = 6
        Case 7
            'Used for bonus twinkles.
            Balls(i).Shape = Int(Rnd(1) * 4) + 7
        Case 9
            'Used for small bonus twinkles.
            Balls(i).Shape = Int(Rnd(1) * 2) + 9
        Case Else
            Balls(i).Shape = Int(Rnd(1) * Size) + 1
      End Select
      Balls(i).Color = Color
      Balls(i).Elastic = AbsorbPct / 100
      'Now fudge out of the loop and initialize the next ball.
      Exit For
    End If
  Next i

Next j

End Sub

Public Function NoBalls() As Boolean

'This function returns a true if no balls are active,
'and a false if even one is bouncing.

Dim i As Integer

For i = 1 To BALLMAX
  If Balls(i).Enabled = True Then
    NoBalls = False
    Exit Function
  End If
Next i

NoBalls = True

End Function
