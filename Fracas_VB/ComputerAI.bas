Attribute VB_Name = "ComputerAI"
Option Explicit
Option Base 1

Dim Score() As Single
Dim BestScore As Single
Dim BestCountry As Long
Dim NumReinforce As Single
Dim AttackDiff As Single
Dim AttackedCountryThisTurn As Long
Dim ToWhere As Long
Dim XferAmount As Single
Public MsgXferAmount As Long   'Public so we can see it for messages.

Dim Cntry As Long
Dim Cntry2 As Long

'These variables are used to calculate the number of troops to move.
Dim AttProp As Single
Dim AttSum1 As Single
Dim AttSum2 As Single
Dim A1 As Single
Dim A2 As Single
Dim A3 As Single
Dim A4 As Single
Dim A5 As Single
Dim A6 As Single
Dim B1 As Single
Dim B2 As Single
Dim B3 As Single
Dim B4 As Single
Dim B5 As Single
Dim B6 As Single
Dim D1 As Single
Dim D2 As Single

'Used for Custom AIs.
Dim AIFile As String
Dim CurrentLine As String
Dim AIFileNumber As Long

'Temporary registers used for a variety of purposes.
Dim TempNum1 As Single
Dim TempNum2 As Single
Dim HQTemp As Single

'The Hate array is used to keep track of, well, how much each computer
'hates each other player.
'For example:  a = Hate(x, y)
'a = the hate value, from 1 to 100.
'x = the player doing the hating.
'y = the player being hated.
Public Hate(6, 6) As Integer

'The following variables define the number of available computer
'personalities to pick from.
Public Const NUMPERSONALITIES As Integer = 12
Public PersonalityName(NUMPERSONALITIES) As String

'These constants define the score for an impossible or unlikely move.
Const NOTPOSSIBLE As Long = -999999   'Just a very low number.
Const UNLIKELY As Long = 1            'Higher than NOTPOSSIBLE, but any other move is better.

'These variables define the personality of the current AI.
Dim SAMESCOREKEEPER         As Single               'Chance of keeping a country with the same bestscore.
Dim HATEFACTOR              As Integer              'How badly this personality holds a grudge.
Dim HQTROOPSMULTIPLIER      As Integer              'Number of times that the troops in a country count.
Dim HQBODYOFWATER           As Integer              'Number of points a body of water is worth to HQ.
Dim HQBORDERINGCOUNTRY      As Integer              'Number of points a bordering country is worth to HQ.
Dim HQONEAWAYMULT           As Single               'Fraction of troops counted for countries 1 away from HQ.
Dim HQTWOAWAYMULT           As Single               'Fraction of troops counted for countries 2 away from HQ.
Dim HQNEXTTOENEMYHQ         As Single               'Score for putting HQ near enemy HQs.
Dim HQNEARCENTERMULT        As Single               'Factor for HQ being closer to the center of the map.
Dim HQDEFENDPROPENSITY      As Integer              'Points per strength difference to defend HQ.
Dim REDEFLOSINGPROPOSITION  As Integer              'Factor for reinforcing defensive troops in a losing battle.
Dim REDEFWINNINGPROPOSITION As Integer              'Factor for reinforcing defensive troops in a winning battle.
Dim REATTLOSINGPROPOSITION  As Integer              'Factor for reinforcing offensive troops in a losing battle.
Dim REATTWINNINGPROPOSITION As Integer              'Factor for reinforcing offensive troops in a winning battle.
Dim REPUTPRESSUREONHQBASE   As Integer              'Score for reinforcing near an enemy HQ.
Dim REPUTPRESSUREONHQMULT   As Integer              'Pro-rated Score for reinforcing near an enemy HQ.
Dim REBETONASURETHING       As Integer              'Score for reinforcing a battle that can already be won.
Dim ACANNEXTROOPS           As Single               'Factor for annexing a country with troops in it.
Dim ACANNEXBASE             As Single               'Score for just annexing any piece of land. (* CFGBonusTroops)
Dim ACANNEXNEARENEMYMULT    As Single               'Multiplier for taking free troops near enemy.
Dim ACANNEXNEXTTOHQ         As Single               'Score for annexing unclaimed land next to an enemy HQ.
Dim ACANNEXTWOAWAYMULT      As Single               'Fraction that free troops two countries away are worth.
Dim ACPORTBASE              As Integer              'Score for just building a port.            (* CFGShips)
Dim ACPORTTROOPS            As Integer              'Factor for building a port on a country with troops.
Dim ACATTACKBASE            As Integer              'Score for simply attacking another defending player.
Dim ACATTACKFACTOR          As Integer              'Pro-rated score for attacking another defending player.
Dim ACBULLYBASE             As Integer              'Score for beating up a player in an easy battle.
Dim ACBULLYFACTOR           As Integer              'Pro-rated score for attacking in an easy battle.
Dim ACCLEANUPEMPTYS         As Integer              'Score for attacking an empty enemy country.
Dim ACATTACKENEMYHQ         As Integer              'Score for being able to attack enemy HQ.
Dim ACHATEATTACKED          As Integer              'How much more I hate you for attacking me.
Dim ACHATEFORGIVEN          As Integer              'How much more I hate you now that I've attacked you.
Dim ACHATEENEMYOFMYENEMY    As Integer              'How much more I hate you now that you've attacked someone else.
Dim TMPUTPRESSUREONHQBASE   As Single               'Score for moving troops near an enemy HQ.
Dim TMPUTPRESSUREONHQMULT   As Single               'Pro-rated score for moving troops near an enemy HQ.
Dim TMSWOOPINSCORE          As Single               'Score for moving troops into a just-taken country.
Dim TMDEFENDPROPENSITY      As Single               'Pro-rated score for defending a country with a troop movement.
Dim TMATTACKPROPENSITY      As Single               'Pro-rated score for attacking a country with a troop movement.
Dim TMBETONASURETHING       As Single               'Score for moving troops to attack a country that can already be taken.
Dim TMMOVEFROMHQ            As Single               'Score for HQ being the only country to get troops from.
Dim TMMAXPCTFROMHQ          As Single               'The maximum percentage of troops that can be moved from HQ.
Dim TMTWOAWAYDEFENSEFACTOR  As Single               'The percent of a normal score for defending a country two moves away.

Public Sub AIchooseHQ(AutoPick As Boolean, Turn As Integer, Phase As Integer)

Dim i As Long
Dim j As Long
Dim k As Long

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

If PlayerType(Turn) = PTYPE_COMPUTER Then
  Call GetPersonalityData(Turn)
Else
  Call GetPersonalityData(7)  'Default to Stonewall if we're autopicking HQ.
End If

'This sub picks the computer's first country.
Call AIsetupscore

'First, we check to see which country has the highest troop total in and
'around it.  If initial troops are at none, then we skip this step.
'The troops in the country itself count more than once.
'Unclaimed bordering troops count once.

DoEvents
For i = 1 To MyMap.NumberOfCountries
  
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  
  'If this country is already owned, skip it for now.
  If MyMap.Owner(i) = 0 Then
    
    'First add the troop score for the country itself.
    Score(i) = Score(i) + (MyMap.TroopCount(i) * HQTROOPSMULTIPLIER)
    
    'Now let's look at this country's neighbors.
    For j = 1 To MyMap.MaxNeighbors
      Cntry = MyMap.Neighbors(i, CInt(j))
      
      'If we hit a zero, then we're done.
      If (Cntry = 0) Then Exit For
      
      If Cntry < TILEVAL_COASTLINE Then   'It's land.
        
        'If this country is not owned, it is worth some points.
        If MyMap.Owner(Cntry) = 0 Then
          Score(i) = Score(i) + (MyMap.TroopCount(Cntry) * HQONEAWAYMULT)  'Score for troops in it.
          Score(i) = Score(i) + HQBORDERINGCOUNTRY  'Score for bordering a country.
          
          'Now we check all of this neighbor's neighbors for more stuff.
          For k = 1 To MyMap.MaxNeighbors
            Cntry2 = MyMap.Neighbors(Cntry, CInt(k))
            If (Cntry2 = 0) Then Exit For
              If Cntry2 < TILEVAL_COASTLINE Then   'It's land.
                
                'Now see if this country already touches the first one!
                If (AIxCanReachy(Cntry2, i) = 0) Then
                 
                 'Add the two-away score for troops.
                  If MyMap.Owner(Cntry2) = 0 Then
                    'Unowned country two steps away.
                    Score(i) = Score(i) + (MyMap.TroopCount(Cntry2) * HQTWOAWAYMULT)
                  Else
                    'It's an enemy HQ.
                    Score(i) = Score(i) + ((HQNEXTTOENEMYHQ * (4 ^ CFGInitTroopCt)) / 2)
                  End If
                End If
              Else   'It's water.
                'Right now, do nothing.
              End If
          Next k
        Else
          'This neighbor is owned by someone else.  Scoring...
          Score(i) = Score(i) + (HQNEXTTOENEMYHQ * (4 ^ CFGInitTroopCt))
        End If
      Else   'It's water.
        Score(i) = Score(i) + HQBODYOFWATER
      End If
    Next j
    
    'Now, let's add a factor for how close to the edge of the map it is.
    'We're using the DisplaySpot to calculate this.  Not perfect, but
    'very elegant!  The closer to the edge, the better.
    'First, check the horizontal direction.
    If MyMap.DigitCoords(i, 1) > (MyMap.Xsize / 2) Then
      TempNum1 = MyMap.Xsize - MyMap.DigitCoords(i, 1)
    Else
      TempNum1 = MyMap.DigitCoords(i, 1)
    End If
    
    'Now the vertical direction.
    If MyMap.DigitCoords(i, 2) > (MyMap.Ysize / 2) Then
      TempNum2 = MyMap.Ysize - MyMap.DigitCoords(i, 2)
    Else
      TempNum2 = MyMap.DigitCoords(i, 2)
    End If
    
    'Now add in scores for each.
    Score(i) = Score(i) + ((TempNum1 + TempNum2) * HQNEARCENTERMULT)
  
  Else
    'This country is already owned.  We *can't* choose it as our country.
    Score(i) = NOTPOSSIBLE
  End If
Next i

Call AIfindbestscore(Turn)

If AutoPick = False Then
  If MyNetworkRole = NW_SERVER Then SendClickToNetwork Turn, Phase, BestCountry, 0
  Call ChooseHQProc(BestCountry, Turn)
Else
  'We're picking this HQ quickly for automatic HQ selection.  No delays.
  Call ClaimHQ(BestCountry, Turn)
End If

End Sub

Public Sub AIreinforce(Turn As Integer, Phase As Integer)

Dim i As Long
Dim j As Long

Call GetPersonalityData(Turn)

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

HQTemp = AIassessHQ(Turn)

NumReinforce = CFGBonusTroops * NumOccupied(Turn)

'This sub finds the best place to dump reinforcement troops.
'The spot is based on which country needs the most defense right now,
'and which country could use the force to attack.

Call AIsetupscore

'First, find which country would benefit most from defensive troops.
DoEvents
For i = 1 To MyMap.NumberOfCountries
  
  
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  
  '-------------------------------------------------------------------------------------------
  If MyMap.Owner(i) = Turn Then
    'We do own this country, so it is a contender...
    If ((MyMap.CountryType(i) And 2) = 2) Then
      'We found our HQ!  Let's move our HQ assessment over.
      Score(i) = Score(i) + HQTemp
    End If
    
    For j = 1 To 6
      'Let's find out what would happen to this country if each other player
      'were to attack it.
      If j <> Turn Then
        
        Call CalculateStrengths(CInt(j), i)
        If AttackStrength > DefendStrength Then
          
          'We could lose this country!
          If (AttackStrength - DefendStrength) > NumReinforce Then
            'Even reinforcing it with these troops won't guarantee safety.
            Score(i) = Score(i) + (((NumReinforce + DefendStrength) / _
                      (NumReinforce + DefendStrength + AttackStrength)) * _
                      REDEFLOSINGPROPOSITION)
          Else
            'We could prevent an invasion with these troops.
            Score(i) = Score(i) + (((NumReinforce + DefendStrength) / _
                      (NumReinforce + DefendStrength + AttackStrength)) * _
                      REDEFWINNINGPROPOSITION)
          End If
        Else
          'This country won't be lost just yet, so we don't need troops here.
        End If
      End If
    Next j
  Else
    'We don't own this country, so we can't exactly put any troops here.
    Score(i) = NOTPOSSIBLE
  End If
  '-------------------------------------------------------------------------------------------
'Next i
'
'DoEvents
'For i = 1 To MyMap.NumberOfCountries
'  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  
  '-------------------------------------------------------------------------------------------
  'Now, find which country would benefit most from offensive troops.
  If MyMap.Owner(i) = Turn Then
    For j = 1 To MyMap.NumberOfCountries
      If (MyMap.Owner(j) <> Turn) And _
            ((AIxCanReachy(i, j) = 1) Or ((AIxCanReachy(i, j) = 2) And _
            ((MyMap.CountryType(i) And 1) = 1))) Then
        
        Call CalculateStrengths(Turn, j)
        AttackDiff = AttackStrength - DefendStrength
        
        'Only look at countries where we have a military influence.
        If AttackStrength <> 0 Then
          If AttackStrength > DefendStrength Then
            'We already can take this country, so we probably don't need troops here.
            Score(i) = Score(i) + REBETONASURETHING
          
          Else
            'We can't take this country yet..
            If (AttackStrength + NumReinforce) > DefendStrength Then
              'We could take this country if we put our troops here!
              Score(i) = Score(i) + (((NumReinforce + AttackStrength) / _
                    (NumReinforce + AttackStrength + DefendStrength)) * _
                    REATTWINNINGPROPOSITION)
            Else
              'We really won't be able to take this country.
              Score(i) = Score(i) + (((NumReinforce + AttackStrength) / _
                    (NumReinforce + AttackStrength + DefendStrength)) * _
                    REATTLOSINGPROPOSITION)
              'But, if it is an enemy HQ, we should put pressure on it!
              If ((MyMap.CountryType(j) And 2) = 2) And (MyMap.Owner(j) <> Turn) Then
                If (AIxCanReachy(i, j)) = 1 Then
                  'We can get to the HQ by land this way.  Full points!
                  Score(i) = Score(i) + REPUTPRESSUREONHQBASE + (((NumReinforce + AttackStrength) / _
                    (NumReinforce + AttackStrength + DefendStrength)) * _
                    REPUTPRESSUREONHQMULT)
                ElseIf (AIxCanReachy(i, j) = 2) Then
                  'We can get to the HQ by water this way.  Fractional points.
                  Score(i) = Score(i) + (ShipPct * (REPUTPRESSUREONHQBASE + (((NumReinforce + AttackStrength) / _
                    (NumReinforce + AttackStrength + DefendStrength)) * _
                    REPUTPRESSUREONHQMULT)))
                End If
              End If
            End If
          End If
        End If
      End If
    Next j
  End If
  '-------------------------------------------------------------------------------------------
'Next i
'
'DoEvents
'For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  
  
  '-------------------------------------------------------------------------------------------
  'Check out port countries.  If there are no enemies on the same shores,
  'then don't bother!
  If (MyMap.Owner(i) = Turn) And ((MyMap.CountryType(i) And 1) = 1) Then
    'This is one of our port countries.
    TempNum2 = 0
    For j = 1 To MyMap.NumberOfCountries
      If (AIxCanReachy(i, j) = 2) And (MyMap.Owner(j) <> Turn) And (MyMap.Owner(j) > 0) Then
        'Here's a country that we could put pressure on by reinforcing this port.
        TempNum2 = TempNum2 + 1
      End If
    Next j
    If TempNum2 = 0 Then
      'Reinforcing this port will do no good.
      Score(i) = UNLIKELY
    End If
  End If
  '-------------------------------------------------------------------------------------------


  '-------------------------------------------------------------------------------------------
  'Now, we need to knock out the countries that are maxed out.  Otherwise, the
  'computer will keep dumping troops into the same country over and over!
  If (MyMap.Owner(i) = Turn) And (MyMap.TroopCount(i) >= MAX_COUNTRY_CAPACITY) Then '(MyMap.TroopCount(i) + (NumReinforce)) >= MAX_COUNTRY_CAPACITY Then
    'We're going to lose our reinforcements.  Let's make this impossible.
    Score(i) = NOTPOSSIBLE
  End If
  '-------------------------------------------------------------------------------------------
Next i

Call AIfindbestscore(Turn)

Call WaitForNoBalls

If BestScore < 0 Then
  'We're passing.
  Call SendPassToNetwork(Turn, Phase, 0)
  Call SNDWhistle
  Exit Sub
End If

If MyNetworkRole = NW_SERVER Then SendClickToNetwork Turn, Phase, BestCountry, 0
Call ReinforceProc(BestCountry, Turn)

End Sub

Public Sub AIaction(Turn As Integer, Phase As Integer)

Dim i As Long
Dim j As Long

Call GetPersonalityData(Turn)

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

'This sub calculates the best course of action this turn by calculating
'scores based on annexing, adding ports, and attacking.

AttackedCountryThisTurn = 0
Call AIsetupscore

DoEvents
For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  
  TempNum1 = AIreachable(i, Turn)
  If TempNum1 = 0 Then
    'If we can't even touch that country, then screw it!
    Score(i) = NOTPOSSIBLE
  
  ElseIf (TempNum1 = -1) And ((MyMap.CountryType(i) And 1) <> 1) And (Coastal(i) = True) Then
    'Let's consider a port here on our land since it doesn't have one.
    'Need to include the Ships setting in our equation.  The more damage a
    'port can do, the better it is to have one!
    Score(i) = Score(i) + (ACPORTBASE * ShipPct) + ((MyMap.TroopCount(i) * ACPORTTROOPS) / (5 - CFGShips))
    'But, a port here may not be necessary.  Let's see if there are
    'any enemy countries at all overseas.
    TempNum2 = 0
    For j = 1 To MyMap.NumberOfCountries
      If (AIxCanReachy(i, j) = 2) And (((MyMap.Owner(j) <> Turn) And (MyMap.Owner(j) > 0)) Or _
               ((MyMap.Owner(j) = 0) And (MyMap.TroopCount(j) > 0))) Then
        'Here's a country that we could put pressure on with a port.
        TempNum2 = TempNum2 + 1
      End If
    Next j
    If TempNum2 = 0 Then
      'A port here will do no good.
      Score(i) = UNLIKELY
    End If
  
  ElseIf (MyMap.Owner(i) = 0) Then
    'Ah, this country can be annexed...but we don't want to take just anything.
    'The number of troops we'll get for it next turn is important, as is the
    'number of troops in it that we'll be inheriting.
    Score(i) = Score(i) + (ACANNEXBASE * (CFGBonusTroops + 0.1)) + (MyMap.TroopCount(i) * ACANNEXTROOPS)
    'The total defense that we can give the new country will be the current
    'attack total for it, plus the troops in it already.
    Call CalculateStrengths(Turn, i)
    TempNum2 = AttackStrength + MyMap.TroopCount(i)
    'Now we will check the attack strength for each other player on this country.
    For j = 1 To 6
      If j <> Turn Then
        Call CalculateStrengths(CInt(j), i)
        If (AttackStrength + (NumOccupied(j) * CFGBonusTroops)) > TempNum2 Then
          'If we take this country, we will need to move troops here this turn
          'or we will lose it.  For now, we'll steer clear.  We add the troopcount
          'here because if there's nothing better to do, at least we'll grab
          'the one with the most troops!
          Score(i) = UNLIKELY + (MyMap.TroopCount(i) * ACANNEXNEARENEMYMULT)
          Exit For
        End If
      End If
    Next j
    'Now, let's look at all the neighbors of this country and see if there
    'are more goodies two countries beyond us.
    For j = 1 To MyMap.MaxNeighbors
      Cntry = MyMap.Neighbors(i, CInt(j))
      If (Cntry = 0) Then Exit For
      If (Cntry < TILEVAL_COASTLINE) Then   'It's land.
        'Now see if this country already borders the first.
        If (AIxCanReachy(Cntry, i) = 0) Then
          'It can't get to the original, so let's add the score.
          If (MyMap.Owner(Cntry) = 0) And (MyMap.TroopCount(Cntry) > 0) Then
            Score(i) = Score(i) + (MyMap.TroopCount(Cntry) * ACANNEXTWOAWAYMULT)
          ElseIf ((MyMap.CountryType(Cntry) And 2) = 2) And _
                                        (MyMap.Owner(Cntry) <> Turn) Then
            'We're two away from an enemy HQ.  Let's tussle.
            Score(i) = Score(i) + ACANNEXNEXTTOHQ
          End If
        End If
      Else    'It's water.
        'Right now, do nothing.
      End If
    Next j
  
  ElseIf (MyMap.Owner(i) <> Turn) And TempNum1 > 0 Then
    'Maybe we can attack this country!
    Call CalculateStrengths(Turn, i)
    If AttackStrength > DefendStrength Then
      'Yes, we can attack it and do damage.
      'There are several factors at work here.  First, the bully factor.
      'The more we can damage it, the more we are interested in attacking it!
      Score(i) = Score(i) + ACBULLYBASE + ((AttackStrength / (AttackStrength + DefendStrength)) * ACBULLYFACTOR)
      'Also, if a country has *no* troops in it, we add in the cleanup factor.
      If MyMap.TroopCount(i) = 0 Then
        Score(i) = UNLIKELY + ACCLEANUPEMPTYS
      End If
      'Second, the cautious factor.  Kill the larger enemy first.
      'The more the opponent can defend, the more we are interested in it.
      Score(i) = Score(i) + ACATTACKBASE + ((DefendStrength / (AttackStrength + DefendStrength)) * ACATTACKFACTOR)
      'Check if we're lookin' at an HQ.
      If ((MyMap.CountryType(i) And 2) = 2) Then
        'Whoa -- this is another player's HQ!  We should put extra effort
        'Into killing it.
        Score(i) = Score(i) + ACATTACKENEMYHQ
      End If
      'Third, the hate factor.  If we hate this person, we are more inclined
      'to attack them over other players.
      Score(i) = Score(i) + (Hate(Turn, MyMap.Owner(i)) * HATEFACTOR)
    Else
      'We can't attack this country at all.
      Score(i) = NOTPOSSIBLE
    End If
  End If
Next i

Call AIfindbestscore(Turn)

Call WaitForNoBalls

If BestScore <= 0 Then
  'We're passing.
  Call SendPassToNetwork(Turn, Phase, 0)
  Call SNDWhistle
  Exit Sub
End If

If MyNetworkRole = NW_SERVER Then SendClickToNetwork Turn, Phase, BestCountry, 0
ActionProc BestCountry, Turn

End Sub

Public Sub AItroopmove(Turn As Integer, Phase As Integer)

Dim i As Long
Dim j As Long
Dim k As Long

Call GetPersonalityData(Turn)

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

'This sub calculates if a troop movement will be beneficial, and how many
'troops to move.  There are three things that might necessitate a troop
'movement:
'1 - to defend a friendly country, especially HQ.
'2 - to attack an enemy country, especially HQ.
'3 - to 'swoop in' after an enemy country has been beaten, and generally
'    make the computer look like it knows what it's doing.   :)

Call AIsetupscore
XferAmount = 0
MsgXferAmount = 0

'Let's see what our HQ assessment is.
HQTemp = AIassessHQ(Turn)

'Do this now in case we pass our turn -- the SNDWhistle would be too fast.
Call WaitForNoBalls

'Situation 1:  Defend a country of our own.
DoEvents
For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  If MyMap.Owner(i) = Turn And MyMap.TroopCount(i) < MAX_COUNTRY_CAPACITY Then
    'This is a valid candidate for a troop recipient.
    If ((MyMap.CountryType(i) And 2) = 2) Then
      'We found our HQ!  Let's move our HQ assessment over and add it in.
      Score(i) = Score(i) + HQTemp
    End If
    For j = 1 To 6
      'Let's find out what would happen to this country if each other player
      'were to attack it.
      If j <> Turn Then
        Call CalculateStrengths(CInt(j), i)
        If (AttackStrength > DefendStrength) Then   'Also prevents /0.
          'We could lose this country!
          Score(i) = Score(i) + (((DefendStrength) / (DefendStrength + AttackStrength)) * _
                                TMDEFENDPROPENSITY)
        Else
          'We're holding this country pretty well right now, so there's no rush to
          'move troops here.
          Score(i) = UNLIKELY
        End If
      End If
    Next j
    'Now let's look at each of this country's neighbors and see if one of
    'them needs some defending.
    For j = 1 To MyMap.NumberOfCountries
      If ((AIxCanReachy(i, j) = 1) Or ((AIxCanReachy(i, j) = 2) And _
            ((MyMap.CountryType(i) And 1) = 1))) Then
        'We found a country that would benefit from us moving troops here.
        If ((MyMap.CountryType(j) And 2) = 2) Then
        'Our HQ is next to this one.  Moving here will defend it.
          Score(i) = Score(i) + (HQTemp * TMTWOAWAYDEFENSEFACTOR)
        End If
        For k = 1 To 6
          'Let's find out what would happen to this two-away country if each
          'other player were to attack it.
          If k <> Turn Then
            Call CalculateStrengths(CInt(k), j)
            If (AttackStrength > DefendStrength) Then   'Also prevents /0.
              'That country is in danger!
              Score(i) = Score(i) + (((DefendStrength) / (DefendStrength + AttackStrength)) * _
                                (TMDEFENDPROPENSITY * TMTWOAWAYDEFENSEFACTOR))
            Else
              'We're holding this country pretty well right now, so there's no rush to
              'move troops here.
              Score(i) = Score(i) + UNLIKELY
            End If
          End If
        Next k
      End If
    Next j
  End If
Next i

'Situation 2:  Attack an enemy country.
For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  If MyMap.Owner(i) = Turn And MyMap.TroopCount(i) < MAX_COUNTRY_CAPACITY Then
    'This is a valid candidate for a troop recipient.
    For j = 1 To MyMap.NumberOfCountries
      If (MyMap.Owner(j) <> Turn) And _
            ((AIxCanReachy(i, j) = 1) Or ((AIxCanReachy(i, j) = 2) And _
            ((MyMap.CountryType(i) And 1) = 1))) Then
        Call CalculateStrengths(Turn, j)
        'Only look at countries where we have a military influence.
        'If AttackStrength <> 0 Then
          If AttackStrength > DefendStrength Then
            'We already can take this country, so we probably don't need troops here.
            Score(i) = Score(i) + TMBETONASURETHING
          Else
            'We can't take this country yet..
            If (AttackStrength + DefendStrength) > 0 Then    'Also prevents /0.
              'We maybe could take this country if we moved troops here!
              Score(i) = Score(i) + ((AttackStrength / _
                    (AttackStrength + DefendStrength)) * TMATTACKPROPENSITY)
            End If
          End If
          'If this is an enemy HQ...
          If ((MyMap.CountryType(j) And 2) = 2) Then
            If (AIxCanReachy(i, j)) = 1 And ((AttackStrength + DefendStrength) > 0) Then
              'We can get to the HQ by land this way.  Full points!
              Score(i) = Score(i) + TMPUTPRESSUREONHQBASE + ((AttackStrength / _
                (AttackStrength + DefendStrength)) * TMPUTPRESSUREONHQMULT)
            ElseIf (AIxCanReachy(i, j) = 2) Then
              'We can get to the HQ by water this way.  Fractional points.
              Score(i) = Score(i) + TMPUTPRESSUREONHQBASE + ((AttackStrength / _
                (AttackStrength + DefendStrength)) * TMPUTPRESSUREONHQMULT * ShipPct)
            End If
          End If
        'End If
      End If
    Next j
  End If
Next i

'Situation 3:  Swoop in after taking a country this turn.
'First, we check to see if we attacked a country this turn, and that we
'did in fact take it over.
If AttackedCountryThisTurn > 0 Then
  'Yup, we attacked one.
  If MyMap.Owner(AttackedCountryThisTurn) = Turn Then
    'Yup, we own it now.  Let's augment the previously calculated scores.
    Score(AttackedCountryThisTurn) = Score(AttackedCountryThisTurn) + TMSWOOPINSCORE
  End If
End If

'Now we knock out the impossibles.
'Let's check if our HQ is the only country we can take troops from.
'Also see if there is *any* adjacent country that can donate.
For i = 1 To MyMap.NumberOfCountries
'  DoEvents
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  If MyMap.Owner(i) = Turn And MyMap.TroopCount(i) < MAX_COUNTRY_CAPACITY Then
    'Let's find all countries of ours that can get here.
    TempNum1 = 0   'Number of adjacent countries that can donate.
    TempNum2 = 0   '1 if our HQ can donate.
    For j = 1 To MyMap.NumberOfCountries
      If MyMap.Owner(j) = Turn And MyMap.TroopCount(j) > 0 And _
         ((AIxCanReachy(j, i) = 1) Or ((AIxCanReachy(j, i) = 2) And _
         ((MyMap.CountryType(j) And 1) = 1))) Then
         'Yup, it can get here.
         TempNum1 = TempNum1 + 1
         'Check if it's our HQ...
         If ((MyMap.CountryType(j) And 2) = 2) Then
           TempNum2 = TempNum2 + 1
         End If
      End If
    Next j
    If (TempNum1 = 0) Then
      'No one to give.  Not good.
      Score(i) = NOTPOSSIBLE
    End If
    If ((TempNum1 = 1) And (TempNum2 = 1)) Then
      'The only one who can give to this country is HQ.  Need to pro-rate this.
      Score(i) = Score(i) + TMMOVEFROMHQ
    End If
  Else
    'We don't own this country, or we're toked on troops.
    'So we can't exactly put any troops here.
    Score(i) = NOTPOSSIBLE
  End If
Next i

Call AIfindbestscore(Turn)

If BestScore <= 0 Then
  'Nobody really needs troops.
  Call SendPassToNetwork(Turn, Phase, 0)
  Call SNDWhistle
  Exit Sub
End If

'Now we know who needs troops the most, but how many, and from where?
'We're assuming that when we get here, we have at least one friendly
'neighbor with troops in it.

'Reset the score array!  Our 'TO' country is BestCountry.
ToWhere = BestCountry
Call AIsetupscore

'Consider the defense of the country we're moving from.
For i = 1 To MyMap.NumberOfCountries
'  DoEvents
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  If MyMap.Owner(i) = Turn And MyMap.TroopCount(i) > 0 And _
         ((AIxCanReachy(i, ToWhere) = 1) Or ((AIxCanReachy(i, ToWhere) = 2) And _
         ((MyMap.CountryType(i) And 1) = 1))) Then
    'We found a friendly country with troops in it who can reach ToWhere.
    For j = 1 To 6
      'Let's find out what would happen to this country if each other player
      'were to attack it.
      If j <> Turn Then
        Call CalculateStrengths(CInt(j), i)
        'Let's add the defense proportion to this country.
        'Basically, if a country is well defended, it gets a high score
        'meaning troops can leave here pretty easily.
        'If a country isn't well defended, it gets a low score and troops
        'are likely to stay (and be reinforced next turn, probably!)
        Score(i) = Score(i) + (((DefendStrength) / (DefendStrength + AttackStrength)) * _
                                TMDEFENDPROPENSITY)
      End If
    Next j
  End If
Next i

'Now that we've gone through all that, we need to knock out countries
'that we don't want to take troops from, like HQ countries.

For i = 1 To MyMap.NumberOfCountries
'  DoEvents
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  If MyMap.Owner(i) = Turn And MyMap.TroopCount(i) > 0 And _
         ((AIxCanReachy(i, ToWhere) = 1) Or ((AIxCanReachy(i, ToWhere) = 2) And _
         ((MyMap.CountryType(i) And 1) = 1))) And ((MyMap.CountryType(i) And 2) = 2) Then
    'It's an HQ.  We need to add the personality constant...
    Score(i) = Score(i) + TMMOVEFROMHQ
  End If
Next i

Call AIfindbestscore(Turn)

If BestScore <= 0 Then
  'Nowhere to get troops from, really.
  Call SendPassToNetwork(Turn, Phase, 0)
  Call SNDWhistle
  Exit Sub
End If

'Now BestCountry is the FROM country, and ToWhere is the TO country.
'We just need to know how much!

'First we get the attack strengths of each player on the TO country.
Call CalculateStrengths(1, ToWhere)
A1 = AttackStrength
Call CalculateStrengths(2, ToWhere)
A2 = AttackStrength
Call CalculateStrengths(3, ToWhere)
A3 = AttackStrength
Call CalculateStrengths(4, ToWhere)
A4 = AttackStrength
Call CalculateStrengths(5, ToWhere)
A5 = AttackStrength
Call CalculateStrengths(6, ToWhere)
A6 = AttackStrength
'Grab the defense of the TO country.
D1 = DefendStrength

'Then we get the attack strengths of each player on the FROM country.
Call CalculateStrengths(1, BestCountry)
B1 = AttackStrength
Call CalculateStrengths(2, BestCountry)
B2 = AttackStrength
Call CalculateStrengths(3, BestCountry)
B3 = AttackStrength
Call CalculateStrengths(4, BestCountry)
B4 = AttackStrength
Call CalculateStrengths(5, BestCountry)
B5 = AttackStrength
Call CalculateStrengths(6, BestCountry)
B6 = AttackStrength
'Grab the defense of the FROM country.
D2 = DefendStrength

'Weigh the overall danger that each country is in.
If (A1 + D1 = 0) Or (A2 + D1 = 0) Or (A3 + D1 = 0) Or (A4 + D1 = 0) Or (A5 + D1 = 0) Or (A6 + D1 = 0) Then
  AttSum1 = 6
Else
  AttSum1 = (A1 / (A1 + D1)) + (A2 / (A2 + D1)) + (A3 / (A3 + D1)) + (A4 / (A4 + D1)) + (A5 / (A5 + D1)) + (A6 / (A6 + D1))
End If

If (B1 + D2 = 0) Or (B2 + D2 = 0) Or (B3 + D2 = 0) Or (B4 + D2 = 0) Or (B5 + D2 = 0) Or (B6 + D2 = 0) Then
  AttSum2 = 6
Else
  AttSum2 = (B1 / (B1 + D2)) + (B2 / (B2 + D2)) + (B3 / (B3 + D2)) + (B4 / (B4 + D2)) + (B5 / (B5 + D2)) + (B6 / (B6 + D2))
End If

'If the quantity below is less than a half, then the FROM country is in more danger.
'Else, the TO country is in more danger, and we should move troops to it!
If (AttSum1 + AttSum2) = 0 Then
  AttProp = 1    'No one attacking either, so why not combine them all?
Else
  AttProp = AttSum1 / (AttSum1 + AttSum2)
End If

If AttProp < 0.5 Then
  'No troops move between these two countries.  Too risky.
  XferAmount = 0
Else
  'We can afford to move a few troops!
  XferAmount = (AttProp - 0.5) * 2 * MyMap.TroopCount(BestCountry)
End If

'Fudge:  if moving troops across water, and both countries have ports,
'and we're not moving from our HQ, why not move them all?
'If AIxcanreachy(BestCountry, ToWhere) = 2 And _
'        ((MyMap.CountryType(BestCountry) And 1) = 1) And _
'        ((MyMap.CountryType(ToWhere) And 1) = 1) And _
'        ((MyMap.CountryType(BestCountry) And 2) <> 2) Then
'  XferAmount = MyMap.TroopCount(BestCountry)
'End If

'Make sure we're not merging two big numbers.
If XferAmount + MyMap.TroopCount(ToWhere) > MAX_COUNTRY_CAPACITY Then
  'Yup, we're being stupid.
  XferAmount = MAX_COUNTRY_CAPACITY - MyMap.TroopCount(ToWhere)
End If

'If we're moving from our HQ, we need to figure in the personality percent.
If ((MyMap.CountryType(BestCountry) And 2) = 2) Then
  XferAmount = XferAmount * TMMAXPCTFROMHQ
End If

'Now, let's move 'em!
If (XferAmount < 1) Or (BestCountry = ToWhere) Or (BestScore <= 0) Then
  'We're passing our turn.
  Call SendPassToNetwork(Turn, Phase, 0)
  Call SNDWhistle
  Exit Sub
End If

'Moose doesn't make troop movements!  We put this check down here for timing reasons.
If (Personality(Turn) = 3) And (PlayerType(Turn) = PTYPE_COMPUTER) Then
  Call SendPassToNetwork(Turn, 4, 0)
  Exit Sub
End If

If MyNetworkRole = NW_SERVER Then SendClickToNetwork Turn, 4, _
       (1000000 * Int(XferAmount)) + (1000 * Int(BestCountry)) + Int(ToWhere), 0
TroopMoveProc BestCountry, ToWhere, XferAmount, Turn

End Sub

Public Sub AIdelay()

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

'This sub just waits a few seconds to give the player time to
'see the computer's moves.

Dim CurrentTime As Date
Dim TargetTime As Date
Dim Delay As Date

If CFGAISpeed = 1 Then Exit Sub

'Otherwise, put a small delay here.
DoEvents     'Update the computer's internal time.

Delay = "00:00:0" & Trim(Str(((CFGAISpeed - 2) * 2) + 1))

CurrentTime = Time
TargetTime = CurrentTime + Delay

Do Until CurrentTime >= TargetTime
  DoEvents
  CurrentTime = Time
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
Loop

End Sub

Public Sub AIsetupscore()

Dim i As Long

'Set up the scores array.
ReDim Score(MyMap.NumberOfCountries)
For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  Score(i) = 0
Next i

End Sub

Public Sub AIfindbestscore(Turn As Integer)

Dim SecondBestScore As Single
Dim i As Long

'This sub finds the country with the best score and puts its number in BestCountry.

BestScore = NOTPOSSIBLE
For i = 1 To MyMap.NumberOfCountries
  If GameMode <> GM_GAME_ACTIVE Then Exit Sub
  'If we beat it or tied it, then we have a new best score.
  If (Score(i) > BestScore) Or ((Score(i) = BestScore) And (Rnd(1) < SAMESCOREKEEPER)) Then
    BestScore = Score(i)
    BestCountry = i
  End If
Next i

'If this is Clyde, he has a large tendency to take the second-best choice.
If (Personality(Turn) = 6) And (Rnd(1) < 0.5) Then
  'Let's find the second-best score.
  SecondBestScore = NOTPOSSIBLE
  For i = 1 To MyMap.NumberOfCountries
    If Score(i) <> BestScore Then
      If (Score(i) > SecondBestScore) Then
        SecondBestScore = Score(i)
        BestCountry = i
      End If
    End If
  Next i
  BestScore = SecondBestScore
End If

End Sub

Public Function AIxCanReachy(FromCountry As Long, ToCountry As Long) As Integer

'This sub checks to see if the FromCountry can reach the ToCountry.
'This sub returns 0 if can't reach, 1 if by land, 2 if by sea, -1 if they are the same.
'Note that we're not considering ports here.

Dim kk As Integer
Dim ll As Integer

'Are they the same country?
If (FromCountry = ToCountry) Then
  AIxCanReachy = -1
  Exit Function
End If

For kk = 1 To MyMap.MaxNeighbors
  If GameMode <> GM_GAME_ACTIVE Then Exit Function
  
  Select Case MyMap.Neighbors(ToCountry, kk)
    Case Is >= TILEVAL_COASTLINE
        For ll = 1 To MyMap.MaxNeighbors
          If MyMap.Neighbors(ToCountry, kk) = MyMap.Neighbors(FromCountry, ll) Then
            'By water.
            AIxCanReachy = 2
            Exit Function
          End If
        Next ll
  
    Case 0
        'Didn't find it.
        AIxCanReachy = 0
        Exit For
  
    Case FromCountry
        'By Land.
        AIxCanReachy = 1
        Exit For
  End Select
Next kk

End Function

Public Function AIreachable(ThisCountry As Long, Turn As Integer)

'This sub checks to see if the passed country is reachable at all by the
'current player.  It returns 0 if can't reach, 1 if by land, 2 if by sea,
'-1 if already owned by this player.
'Note that ports are considered here.

Dim k As Integer
Dim l As Long
Dim m As Integer

If MyMap.Owner(ThisCountry) = Turn Then
  'Already owned.
  AIreachable = -1
  Exit Function
End If

If GameMode <> GM_GAME_ACTIVE Then Exit Function
For k = 1 To MyMap.MaxNeighbors
  
  If MyMap.Neighbors(ThisCountry, k) >= TILEVAL_COASTLINE Then
    For l = 1 To MyMap.NumberOfCountries
      For m = 1 To MyMap.MaxNeighbors
        If (MyMap.Neighbors(ThisCountry, k) = MyMap.Neighbors(l, m)) And _
             (l <> ThisCountry) And (MyMap.Owner(l) = Turn) And _
             ((MyMap.CountryType(l) And 1) = 1) Then
          'By water.
          AIreachable = 2
          Exit Function
        End If
      Next m
    Next l
  
  ElseIf MyMap.Neighbors(ThisCountry, k) = 0 Then
    'Didn't find it.
    AIreachable = 0
    Exit Function
  
  ElseIf MyMap.Owner(MyMap.Neighbors(ThisCountry, k)) = Turn Then
    'By Land.
    AIreachable = 1
    Exit Function
  End If
Next k

End Function

Public Sub SelectionBall(ThisCountry As Long, Size As Integer, Turn As Integer)

'This sub bounces a single ball on the country that the computer just chose.
'Just so the player can see what's happening.

     'Call BuildBalls(1, _
                    (MyMap.DigitCoords(ThisCountry, 1) - 1) * 8, _
                    (MyMap.DigitCoords(ThisCountry, 2) - 1) * 8, _
                    -1, _
                    0, _
                    0, _
                    Size, _
                    Player(Turn))

End Sub

Public Function AIassessHQ(Turn As Integer) As Long

Dim i As Long
Dim j As Long

TempNum2 = 0

If GameMode <> GM_GAME_ACTIVE Then Exit Function
For i = 1 To MyMap.NumberOfCountries
  If (MyMap.Owner(i) = Turn) And ((MyMap.CountryType(i) And 2) = 2) Then
    'Found this player's HQ.
    'Now let's see what each other attacker can do.
    For j = 1 To 6
      If j <> Turn Then
        Call CalculateStrengths(CInt(j), i)
        If (AttackStrength = 0) Or (AttackStrength + DefendStrength = 0) Then   'Avoid /0.
          'This player can't get to our HQ.  Don't add anything.
        Else
          'Someone has attack points on our HQ, so let's add a fraction
          'of the response factor.
          TempNum2 = TempNum2 + (HQDEFENDPROPENSITY * (AttackStrength / (DefendStrength + AttackStrength)))
        End If
      End If
    Next j
  End If
Next i

AIassessHQ = TempNum2

End Function

Public Sub UpdateHate(Attacker As Integer, Attackee As Integer)

Dim i As Long
Dim j As Long

'First, the one who is attacked hates the attacker even more.
Hate(Attackee, Attacker) = Hate(Attackee, Attacker) + ACHATEATTACKED

'Now the attacker relieves some of the pent-up emotion.
Hate(Attacker, Attackee) = Hate(Attacker, Attackee) + ACHATEFORGIVEN

'Each other player likes the attacker that much more.
For i = 1 To 6
  If (i <> Attacker) And (i <> Attackee) Then
    Hate(i, Attacker) = Hate(i, Attacker) + ACHATEENEMYOFMYENEMY
  End If
Next i

'Make sure we're within our hate limits.
For i = 1 To 6
  For j = 1 To 6
    If Hate(i, j) < 1 Then Hate(i, j) = 1
    If Hate(i, j) > 100 Then Hate(i, j) = 100
  Next j
Next i

End Sub

Public Sub WaitForNoBalls()

Dim NoBallsLeft As Boolean

NoBallsLeft = False
Do Until NoBallsLeft Or GameEnding
  NoBallsLeft = NoBalls      'This function is in Boom.bas.
  DoEvents
  If (GameMode <> GM_GAME_ACTIVE) And (GameMode <> GM_DIALOG_OPEN) Then Exit Sub
Loop

End Sub

Public Sub InitPersonalities()

Dim i As Long

For i = 1 To 6
  'Each computer player defaults to Stonewall.
  Personality(i) = 1
Next i

PersonalityName(1) = "Stonewall"
PersonalityName(2) = "Bully"
PersonalityName(3) = "Moose"
PersonalityName(4) = "Ahab"
PersonalityName(5) = "Paranoid"
PersonalityName(6) = "Clyde"
PersonalityName(7) = "Custom 1"
PersonalityName(8) = "Custom 2"
PersonalityName(9) = "Custom 3"
PersonalityName(10) = "Custom 4"
PersonalityName(11) = "Custom 5"
PersonalityName(12) = "Custom 6"

End Sub

Public Sub GetPersonalityData(Turn As Integer)

Dim MyPerson As Integer

If Turn = 7 Then
  MyPerson = 7
Else
  MyPerson = Personality(Turn)
End If

Select Case MyPerson

Case 1, 7:  'Stonewall.
'Probably the best all-around computer player.
  Call Stonewall

Case 2:   'Bully.
'Blindly charges into battles with all it's got.  Picks on smaller countries.
SAMESCOREKEEPER = 0.5         'Chance of keeping a country with the same bestscore.
HATEFACTOR = 40               'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 4        'Number of times that the troops in a country count.
HQBODYOFWATER = -2            'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = 5        'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 1.5           'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.5           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -2          'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.1       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 100      'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 30   'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 50  'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 550  'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 800 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 4100  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 3500  'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 500       'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 2             'Factor for annexing a country with troops in it.
ACANNEXBASE = 10              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 3      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 330         'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 0.8      'Fraction that free troops two countries away are worth.
ACPORTBASE = 8                'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 1              'Factor for building a port on a country with troops.
ACATTACKBASE = 600            'Score for simply attacking another defending player.
ACATTACKFACTOR = 2000         'Pro-rated score for attacking another defending player.
ACBULLYBASE = 1000            'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1000          'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 1000        'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 2900        'Score for being able to attack enemy HQ.
ACHATEATTACKED = 50           'How much more I hate you for attacking me.
ACHATEFORGIVEN = -2           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -5     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 15000 'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 7000  'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 3700         'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 300      'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 7800     'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 3500      'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = 0              'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 1            'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 0.3  'The percent of a normal score for defending a country two moves away.

Case 3:  'Moose.
'Hates troop movements.  Digs in and doesn't budge.
SAMESCOREKEEPER = 0.1         'Chance of keeping a country with the same bestscore.
HATEFACTOR = 10               'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 3        'Number of times that the troops in a country count.
HQBODYOFWATER = -10           'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = -3       'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 1             'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.1           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -11         'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.2       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 700      'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 80   'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 150 'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 50   'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 150 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 1100  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 800   'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 50        'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 3             'Factor for annexing a country with troops in it.
ACANNEXBASE = 12              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 1      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 21          'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 0.9      'Fraction that free troops two countries away are worth.
ACPORTBASE = 10               'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 1              'Factor for building a port on a country with troops.
ACATTACKBASE = 200            'Score for simply attacking another defending player.
ACATTACKFACTOR = 800          'Pro-rated score for attacking another defending player.
ACBULLYBASE = 10              'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1             'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 13          'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 1100        'Score for being able to attack enemy HQ.
ACHATEATTACKED = 9            'How much more I hate you for attacking me.
ACHATEFORGIVEN = -7           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -9     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 0     'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 0     'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 0            'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 0        'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 0        'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 0         'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = 0              'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 0            'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 0    'The percent of a normal score for defending a country two moves away.

Case 4:  'Ahab.
'Loves water.  Always builds toward coastlines.
SAMESCOREKEEPER = 0.7         'Chance of keeping a country with the same bestscore.
HATEFACTOR = 21               'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 3        'Number of times that the troops in a country count.
HQBODYOFWATER = 15            'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = -2       'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 0.5           'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.1           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -10         'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.3       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 900      'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 50   'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 200 'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 300  'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 350 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 1800  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 1200  'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 300       'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 2             'Factor for annexing a country with troops in it.
ACANNEXBASE = 10              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 2      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 63          'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 1        'Fraction that free troops two countries away are worth.
ACPORTBASE = 79               'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 4              'Factor for building a port on a country with troops.
ACATTACKBASE = 200            'Score for simply attacking another defending player.
ACATTACKFACTOR = 1000         'Pro-rated score for attacking another defending player.
ACBULLYBASE = 10              'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1             'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 30          'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 800         'Score for being able to attack enemy HQ.
ACHATEATTACKED = 21           'How much more I hate you for attacking me.
ACHATEFORGIVEN = -8           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -1     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 8000  'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 3000  'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 0            'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 2500     'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 2500     'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 500       'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = -700           'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 0.5          'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 0.6  'The percent of a normal score for defending a country two moves away.

Case 5:  'Paranoid.
'Defensive to the extreme!
SAMESCOREKEEPER = 0.001       'Chance of keeping a country with the same bestscore.
HATEFACTOR = 1                'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 2        'Number of times that the troops in a country count.
HQBODYOFWATER = -15           'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = -2       'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 0.6           'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.2           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -40         'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.8       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 1600     'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 300  'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 350 'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 50   'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 100 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 1500  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 1100  'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 600       'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 2             'Factor for annexing a country with troops in it.
ACANNEXBASE = 10              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 2      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 10          'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 1        'Fraction that free troops two countries away are worth.
ACPORTBASE = 5                'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 2              'Factor for building a port on a country with troops.
ACATTACKBASE = 200            'Score for simply attacking another defending player.
ACATTACKFACTOR = 800          'Pro-rated score for attacking another defending player.
ACBULLYBASE = 10              'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1             'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 30          'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 500         'Score for being able to attack enemy HQ.
ACHATEATTACKED = 3            'How much more I hate you for attacking me.
ACHATEFORGIVEN = -3           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -3     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 1000  'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 500   'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 0            'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 6000     'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 1500     'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 500       'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = -900           'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 0            'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 1    'The percent of a normal score for defending a country two moves away.

Case 6:  'Clyde.
'The 'easy' computer opponent.  Does the occasional dumb move.
'This is handled in code -- the stats are the same as Stonewall.
SAMESCOREKEEPER = 0.2         'Chance of keeping a country with the same bestscore.
HATEFACTOR = 15               'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 3        'Number of times that the troops in a country count.
HQBODYOFWATER = -5            'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = 1        'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 1             'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.3           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -10         'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.1       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 500      'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 50   'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 200 'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 250  'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 300 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 2100  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 1500  'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 50        'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 1             'Factor for annexing a country with troops in it.
ACANNEXBASE = 10              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 2      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 47          'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 0.8      'Fraction that free troops two countries away are worth.
ACPORTBASE = 10               'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 1              'Factor for building a port on a country with troops.
ACATTACKBASE = 200            'Score for simply attacking another defending player.
ACATTACKFACTOR = 1000         'Pro-rated score for attacking another defending player.
ACBULLYBASE = 10              'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1             'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 9           'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 900         'Score for being able to attack enemy HQ.
ACHATEATTACKED = 11           'How much more I hate you for attacking me.
ACHATEFORGIVEN = -3           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -7     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 10000 'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 4000  'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 1500         'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 1100     'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 3800     'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 500       'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = -100           'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 1            'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 0.75 'The percent of a normal score for defending a country two moves away.

Case Else

'We are using a custom personality.  Let's grab the appropriate file and
'read the personality parameters.

AIFile = App.Path & "\Custom" & Trim(Str(Personality(Turn) - 6)) & ".AI"
AIFileNumber = FreeFile

'If the file doesn't exist, we exit.
If Dir(AIFile) = "" Then
  'Default to Stonewall.
  Call Stonewall
  Exit Sub
End If

Open AIFile For Input As AIFileNumber

Input #AIFileNumber, CurrentLine   'Comment Line.

Input #AIFileNumber, CurrentLine: SAMESCOREKEEPER = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HATEFACTOR = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQTROOPSMULTIPLIER = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQBODYOFWATER = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQBORDERINGCOUNTRY = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQONEAWAYMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQTWOAWAYMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQNEXTTOENEMYHQ = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQNEARCENTERMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: HQDEFENDPROPENSITY = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REDEFLOSINGPROPOSITION = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REDEFWINNINGPROPOSITION = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REATTLOSINGPROPOSITION = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REATTWINNINGPROPOSITION = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REPUTPRESSUREONHQBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REPUTPRESSUREONHQMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: REBETONASURETHING = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACANNEXTROOPS = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACANNEXBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACANNEXNEARENEMYMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACANNEXNEXTTOHQ = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACANNEXTWOAWAYMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACPORTBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACPORTTROOPS = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACATTACKBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACATTACKFACTOR = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACBULLYBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACBULLYFACTOR = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACCLEANUPEMPTYS = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACATTACKENEMYHQ = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACHATEATTACKED = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACHATEFORGIVEN = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: ACHATEENEMYOFMYENEMY = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMPUTPRESSUREONHQBASE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMPUTPRESSUREONHQMULT = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMSWOOPINSCORE = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMDEFENDPROPENSITY = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMATTACKPROPENSITY = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMBETONASURETHING = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMMOVEFROMHQ = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMMAXPCTFROMHQ = CutNumberOut(CurrentLine)
Input #AIFileNumber, CurrentLine: TMTWOAWAYDEFENSEFACTOR = CutNumberOut(CurrentLine)

Close #AIFileNumber

End Select
End Sub

Public Function CutNumberOut(CurrentLine As String)

Dim Pos As Integer

On Error GoTo ErrorHandler

'Each line has the form
'VARNAME = XX      'Comment
'The comment may not exist.  We need to grab everything after the
'equals sign, but before an apostrophe (or the end of the line).

'Step 1:  Lop off the equals sign and everything in front of it.
Pos = InStr(CurrentLine, "=")
If (Pos > 0) Then
  'Found the equals sign.
  CurrentLine = Right(CurrentLine, Len(CurrentLine) - Pos)
  'Step 2:  Now find the first apostrophe.
  Pos = InStr(CurrentLine, "'")
  If (Pos > 0) Then
    'Found the apostrophe.  Now save everything before it.
    CurrentLine = Left(CurrentLine, Pos - 1)
    CutNumberOut = Val(Trim(CurrentLine))
  Else
    'Couldn't find an apostrophe!  We'll assume there is no comment.
    CutNumberOut = Val(Trim(CurrentLine))
  End If
Else
  'Couldn't find an equals sign on the line!  Error!
  CutNumberOut = 0
End If
  
Exit Function
  
ErrorHandler:

CutNumberOut = 0
  
End Function

Public Sub Stonewall()

SAMESCOREKEEPER = 0.2         'Chance of keeping a country with the same bestscore.
HATEFACTOR = 15               'How badly this personality holds a grudge.
HQTROOPSMULTIPLIER = 3        'Number of times that the troops in a country count.
HQBODYOFWATER = -5            'Number of points a body of water is worth to HQ.
HQBORDERINGCOUNTRY = 1        'Number of points a bordering country is worth to HQ.
HQONEAWAYMULT = 1             'Fraction of troops counted for countries 1 away from HQ.
HQTWOAWAYMULT = 0.3           'Fraction of troops counted for countries 2 away from HQ.
HQNEXTTOENEMYHQ = -10         'Score for putting HQ near enemy HQs.
HQNEARCENTERMULT = -0.1       'Factor for HQ being closer to the center of the map.
HQDEFENDPROPENSITY = 500      'Points per strength difference to defend HQ.
REDEFLOSINGPROPOSITION = 50   'Factor for reinforcing defensive troops in a losing battle.
REDEFWINNINGPROPOSITION = 200 'Factor for reinforcing defensive troops in a winning battle.
REATTLOSINGPROPOSITION = 250  'Factor for reinforcing offensive troops in a losing battle.
REATTWINNINGPROPOSITION = 300 'Factor for reinforcing offensive troops in a winning battle.
REPUTPRESSUREONHQBASE = 2100  'Score for reinforcing near an enemy HQ.
REPUTPRESSUREONHQMULT = 1500  'Pro-rated Score for reinforcing near an enemy HQ.
REBETONASURETHING = 50        'Score for reinforcing a battle that can already be won.
ACANNEXTROOPS = 1             'Factor for annexing a country with troops in it.
ACANNEXBASE = 10              'Score for just annexing any piece of land. (* CFGBonusTroops)
ACANNEXNEARENEMYMULT = 2      'Multiplier for taking free troops near enemy.
ACANNEXNEXTTOHQ = 47          'Score for annexing unclaimed land next to an enemy HQ.
ACANNEXTWOAWAYMULT = 0.8      'Fraction that free troops two countries away are worth.
ACPORTBASE = 10               'Score for just building a port.            (* CFGShips)
ACPORTTROOPS = 1              'Factor for building a port on a country with troops.
ACATTACKBASE = 200            'Score for simply attacking another defending player.
ACATTACKFACTOR = 1000         'Pro-rated score for attacking another defending player.
ACBULLYBASE = 10              'Score for beating up a player in an easy battle.
ACBULLYFACTOR = 1             'Pro-rated score for attacking in an easy battle.
ACCLEANUPEMPTYS = 9           'Score for attacking an empty enemy country.
ACATTACKENEMYHQ = 900         'Score for being able to attack enemy HQ.
ACHATEATTACKED = 11           'How much more I hate you for attacking me.
ACHATEFORGIVEN = -3           'How much more I hate you now that I've attacked you.
ACHATEENEMYOFMYENEMY = -7     'How much more I hate you now that you've attacked someone else.
TMPUTPRESSUREONHQBASE = 10000 'Score for moving troops near an enemy HQ.
TMPUTPRESSUREONHQMULT = 4000  'Pro-rated score for moving troops near an enemy HQ.
TMSWOOPINSCORE = 1500         'Score for moving troops into a just-taken country.
TMDEFENDPROPENSITY = 1100     'Pro-rated score for defending a country with a troop movement.
TMATTACKPROPENSITY = 3800     'Pro-rated score for attacking a country with a troop movement.
TMBETONASURETHING = 500       'Score for moving troops to attack a country that can already be taken.
TMMOVEFROMHQ = -100           'Score for HQ being the only country to get troops from.
TMMAXPCTFROMHQ = 0.2          'The maximum percentage of troops that can be moved from HQ.
TMTWOAWAYDEFENSEFACTOR = 0.75 'The percent of a normal score for defending a country two moves away.

End Sub

Public Sub ChooseHQProc(BestCountry As Long, Turn As Integer)

Call SelectionBall(BestCountry, 5, Turn)
Call WaitForNoBalls
Call ClaimHQ(BestCountry, Turn)
Call DrawMap
Call SNDPlayFanfare(Player(Turn))

End Sub

Public Sub ReinforceProc(BestCountry As Long, Turn As Integer)

If WonTurn = 1 And WonGame = True Then Exit Sub

Call WaitForNoBalls
Call SelectionBall(BestCountry, 5, Turn)
Call WaitForNoBalls
Call Reinforce(BestCountry, Turn)
Call ShortFlash(BestCountry)
Call DrawMap
Call SNDTroopsIn(Player(Turn))

End Sub

Public Sub ActionProc(BestCountry As Long, Turn As Integer)

If WonTurn = 1 And WonGame = True Then Exit Sub

Call WaitForNoBalls
If MyMap.Owner(BestCountry) = 0 Then
  'We're annexing.
  Call SelectionBall(BestCountry, 5, Turn)
  Call WaitForNoBalls
  Call AnnexCountry(BestCountry, Turn)
  Call ShortFlash(BestCountry)
  Call DrawMap
  Call SNDPlayFanfare(Player(Turn))
ElseIf MyMap.Owner(BestCountry) = Turn Then
  'We're making a port.
  Call SelectionBall(BestCountry, 5, Turn)
  Call WaitForNoBalls
  Call MakePort(BestCountry)
  Call ShortFlash(BestCountry)
  Call DrawMap
  Call SNDBuildAPort
ElseIf MyMap.Owner(BestCountry) <> Turn Then
  'We're attacking.
  Call UpdateHate(CInt(Turn), CInt(MyMap.Owner(BestCountry)))
  Call CalculateStrengths(Turn, BestCountry)   'Gotta get the right numbers again!
  Call AttackCountry(BestCountry, Turn)
  Call ShortFlash(BestCountry)
  Call DrawMap
  AttackedCountryThisTurn = BestCountry        'Used during troop movement.
End If

End Sub

Public Sub TroopMoveProc(BestCountry As Long, ToWhere As Long, XferAmount As Single, Turn As Integer)

If WonTurn = 1 And WonGame = True Then Exit Sub

MsgXferAmount = Int(XferAmount)  'Puts up the right message.
Land.UpdateMessages
DoEvents

Call WaitForNoBalls
Call KillTroops(BestCountry, Int(XferAmount))
Call ShortFlash(BestCountry)
Call DrawMap
Call SelectionBall(ToWhere, 5, Turn)
Call SelectionBall(BestCountry, 6, Turn)
Call SNDTroopsOut(Player(Turn))
Call WaitForNoBalls
Call AddTroops(ToWhere, Int(XferAmount))
Call ShortFlash(ToWhere)
Call DrawMap
Call SNDTroopsIn(Player(Turn))

End Sub

