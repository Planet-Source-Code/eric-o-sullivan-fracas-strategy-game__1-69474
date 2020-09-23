Attribute VB_Name = "Play"
Option Explicit
Option Base 1

Public Sub NextTurn(Turn As Integer, Phase As Integer)

'This sub just bumps up the turn counter.

DoEvents
Land.Oopsie.Caption = ""
If GameMode <> GM_GAME_ACTIVE Then Exit Sub

Call Land.SetupIndicators

Turn = Turn + 1
Phase = 1
If Turn > 6 Then
  Turn = 1
  TurnCounter = TurnCounter + 1
End If

'Set this false so that we'll check for an event at the start of the next
'player's turn.
CheckedForEvent = False
'Clear a few more flags for 2.0 BETA incidents...
AlreadyAddedTroopsThisPhase = False
AlreadyPerformedActionThisPhase = False
AlreadyMovedTroopsThisPhase = False
AlreadyChoseTroopsThisPhase = False
AlreadyCheckedForMaxedOutCountries = False

'Only enable the save game option if this is the start of a human's turn,
'and if the map has been saved already!  No saved games during network games.
If (PlayerType(Turn) = PTYPE_HUMAN) And (NumOccupied(Turn) > 0) And (MyNetworkRole = NW_NONE) Then
  Land.MenuSaveGame.Enabled = True
Else
  Land.MenuSaveGame.Enabled = False
End If

'Check for a random event at the start of this player's turn.
If (CFGEvents > 1) And (CheckedForEvent = False) And (NumOccupied(Turn) > 0) Then     'Only check if we're configured to do so.
  Call RollForRandomEvent(Turn)
  CheckedForEvent = True
End If

Land.SetTurn Turn
Land.SetPhase Phase

Land.SetUpPlayerControls

Land.UpdateMessages
End Sub

Public Sub NextPhase(Turn As Integer, Phase As Integer)

Dim TempNum As Long

DoEvents
Land.Oopsie.Caption = ""
If GameMode <> GM_GAME_ACTIVE Then Exit Sub

Call Land.SetupIndicators

AlreadyPassedThisPhase = False

If Phase = 1 Then
  Phase = 2
  Land.SetPhase Phase
  Land.SetUpPlayerControls
  Land.UpdateMessages
  DoEvents
  Exit Sub
End If

If Phase = 2 Then
  'Let's check and see if we have more than one country.
  'If not, there is no troop movement phase.
  If NumOccupied(Turn) <= 1 Then
    Call NextTurn(Turn, Phase)
    Land.SetUpPlayerControls
    Exit Sub
  End If
  'Let's go to the troop movement phase.
  Land.TroopMoveNum.Text = "0"
  Phase = 3
  Land.SetPhase Phase
  Land.SetUpPlayerControls
  Land.UpdateMessages
  DoEvents
  Exit Sub
End If

If Phase = 3 Then
  'Set up the default number of troops in the text box.
  TempNum = Int(MyMap.TroopCount(TroopMoveSrc))   'Default to all troops.
  If TempNum = 0 Then TempNum = 1
  Land.TroopMoveNum.Text = Trim(Str(TempNum))
  Phase = 4
  Land.SetPhase Phase
  Land.SetUpPlayerControls
  Exit Sub
End If

If Phase = 4 Then
  Call NextTurn(Turn, Phase)
  Land.SetUpPlayerControls
  Exit Sub
End If

End Sub

Public Sub PlayerClick(CurrentMouse As Long, Turn As Integer, Phase As Integer)

'In this sub we perform whatever action the current player
'is trying to do.  We only get here on a human's turn.

'This sub acts on a left mouse click.  First we determine
'which phase of the player's turn it is, then act accordingly.
'Note that CurrentMouse contains the ID of the entity that
'is currently selected.

Dim a As Long
Dim B As Long
Dim i As Integer
Dim j As Integer

Dim TempNum As Long

'If the user clicked on water, then leave this sub.
If CurrentMouse >= TILEVAL_COASTLINE Then Exit Sub

'If the user owns no countries, then Phase isn't important --
'this is their first turn.
If NumOccupied(Turn) = 0 Then
  'Leave if the computer should be picking HQ for us (automatic HQ selection).
  If CFGHQSelect = 2 Then Exit Sub
  'Let's see if the country they chose is already occupied.
  If MyMap.Owner(CurrentMouse) = 0 Then
    'This country is not owned by anyone -- let's give it to this player.
    'This will be this players HQ from now on, so we'll mark it as such.
    Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
    Call ClaimHQ(CurrentMouse, Turn)
    Call DrawMap
    Call SNDPlayFanfare(Player(Turn))
    Call NextTurn(Turn, Phase)
    Exit Sub
  Else
    'This country is already owned.
    Land.Oopsie.Caption = "That country is already taken."
    Call SNDBooBoo
    Exit Sub
  End If
End If

Select Case Phase
Case 1:
  'This is phase 1 -- Place troops.
  'Bugfix for 2.0: If we already added troops this phase, leave.
  'In 1.1 you could add troops forever by clicking repeatedly.
  If AlreadyAddedTroopsThisPhase = True Then Exit Sub
  'If there are no troops to add, then leave!
  If CFGBonusTroops = 0 Then Exit Sub
  If MyMap.Owner(CurrentMouse) = Turn Then
    'The current player owns this country.  Good!
    If MyMap.TroopCount(CurrentMouse) >= MAX_COUNTRY_CAPACITY Then
      Land.Oopsie.Caption = "This country is already maxed out.  The maximum allowable population of any country is 999."
      Call SNDBooBoo
      Exit Sub
    End If
    AlreadyAddedTroopsThisPhase = True
    Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
    Call Reinforce(CurrentMouse, Turn)
    Call SNDTroopsIn(Player(Turn))
    Call ShortFlash(CurrentMouse)
    DrawMap
    Call NextPhase(Turn, Phase)
    Exit Sub
  Else
    'This country is not owned or is owned by someone else.
    Land.Oopsie.Caption = "You must put troops in your own country."
    Call SNDBooBoo
    Exit Sub
  End If

Case 2:
  'This is phase 2 -- attack, annex, or build port.
  'Bugfix for 2.0: If we already performed an action this phase, leave.
  'In 1.1 you could annex multiple countries by clicking repeatedly.
  If AlreadyPerformedActionThisPhase = True Then Exit Sub
  'CurrentMouse is the country we want to attack,annex, or build a port on.
  If MyMap.Owner(CurrentMouse) = 0 Then
    'This country is unowned, so we will annex it -- But
    'only if it is adjacent to a country we own!
    For i = 1 To MyMap.MaxNeighbors
      'Let's look at each neighbor of the chosen country.
      If MyMap.Neighbors(CurrentMouse, i) < TILEVAL_COASTLINE And MyMap.Neighbors(CurrentMouse, i) > 0 Then   'No water!
        'This neighbor is land.
        If MyMap.Owner(MyMap.Neighbors(CurrentMouse, i)) = Turn Then
          'Yup, there's an adjacent country owned by this player.
          AlreadyPerformedActionThisPhase = True
          Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
          Call AnnexCountry(CurrentMouse, Turn)
          Call SNDPlayFanfare(Player(Turn))
          Call ShortFlash(CurrentMouse)
          DrawMap
          Call NextPhase(Turn, Phase)
          Exit Sub
        End If
      ElseIf MyMap.Neighbors(CurrentMouse, i) >= TILEVAL_COASTLINE Then
        'Let's see if the chosen country is accessible by water.
        For a = 1 To MyMap.NumberOfCountries
          'The country we're looking at has to belong to this player,
          'and it has to have a port!
          If (MyMap.Owner(a) = Turn) And ((MyMap.CountryType(a) And 1) = 1) Then
            For j = 1 To MyMap.MaxNeighbors
              'If they share the same water mass as a neighbor...
              If MyMap.Neighbors(a, j) = MyMap.Neighbors(CurrentMouse, i) Then
                'Then they can sail a boat there!  Claim it.
                AlreadyPerformedActionThisPhase = True
                Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
                Call AnnexCountry(CurrentMouse, Turn)
                Call SNDPlayFanfare(Player(Turn))
                Call ShortFlash(CurrentMouse)
                DrawMap
                Call NextPhase(Turn, Phase)
                Exit Sub
              End If
            Next j
          End If
        Next a
      End If
    Next i
    'No adjacent countries are owned by this player.  Ignore the click.
    Land.Oopsie.Caption = "You cannot reach that country to annex it."
    Call SNDBooBoo
    Exit Sub
  
  ElseIf MyMap.Owner(CurrentMouse) <> Turn Then
    'We are attacking someone!
    'Let's calculate our attack strength.  This is the sum of all of this
    'player's troop counts in countries that border this one.
    'We also calculate the defense strength of this country.  This is the
    'sum of all of this country's owner's troop counts in this country and
    'adjacent ones.
    
    Call CalculateStrengths(Turn, CurrentMouse)
        
    If AttackStrength = 0 Then
      'They can't attack this country -- either too far away or no manpower.
      'Leave quietly so they can choose again.
      Land.Oopsie.Caption = "You cannot reach that country to attack it."
      Call SNDBooBoo
      Exit Sub
    End If
  
    'Now we kill stuff!
    
    'When attacking another country, you only need to have more
    'attack strength than your opponent to do damage.  The opponent
    'loses the difference in strengths.  If they drop to zero, this country
    'is ours!  That's all there is to it.
    
    If AttackStrength > DefendStrength Then
      AlreadyPerformedActionThisPhase = True
      Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
      Call AttackCountry(CurrentMouse, Turn)
      Call ShortFlash(CurrentMouse)
      DrawMap
      Call NextPhase(Turn, Phase)
      Exit Sub
    Else
    'Can't attack -- not enough strength!
      Land.Oopsie.Caption = "You do not yet have enough strength to attack that country."
      Call SNDBooBoo
      Exit Sub
    End If
  
  Else
    'This country is owned by this player.  Let's build a port on it!
    'First we see if this country borders water.
    If Coastal(CurrentMouse) = False Then
      'This country doesn't border water!  No point in putting a port on it.
      Land.Oopsie.Caption = "To build a port, a country must border a body of water."
      Call SNDBooBoo
      Exit Sub
    End If
    
    'Does this country already have a port?  Check bit 1.
    If (MyMap.CountryType(CurrentMouse) And 1) = 1 Then
      'We already have one here!  Leave quietly.
      Land.Oopsie.Caption = "That country already has a port."
      Call SNDBooBoo
      Exit Sub
    End If
     
    'Let's put a port here!  Turn on bit 1.
    AlreadyPerformedActionThisPhase = True
    Call SendClickToNetwork(Turn, Phase, CurrentMouse, 0)
    Call MakePort(CurrentMouse)
    Call SNDBuildAPort
    Call ShortFlash(CurrentMouse)
    DrawMap
    Call NextPhase(Turn, Phase)
    Exit Sub
  End If
  
Case 3:
  'This is phase 3 -- Troop movement source select.
  If AlreadyChoseTroopsThisPhase = True Then Exit Sub
  TroopMoveSrc = 0  'To prevent a second click before this is done!
  If MyMap.Owner(CurrentMouse) <> Turn Then
    'We didn't click on one of our countries.
    Land.Oopsie.Caption = "You must move troops from your own country."
    Call SNDBooBoo
    Exit Sub
  ElseIf MyMap.TroopCount(CurrentMouse) = 0 Then
    'This country is ours, but no troops are in it!
    Land.Oopsie.Caption = "There are no troops in that country."
    Call SNDBooBoo
  Else
    'We picked a country of our own that has troops in it.
    AlreadyChoseTroopsThisPhase = True
    Call SNDTroopsOut(Player(Turn))
    Call ShortFlash(CurrentMouse)
    DrawMap
    TroopMoveSrc = CurrentMouse
    Call NextPhase(Turn, Phase)
  End If
  
Case 4:
  'This is phase 4 -- Troop movement destination select.
  'Prevention for 2.0: If we already moved troops this phase, leave.
  'If we aren't yet ready for troop movements, then leave.  (Timing problem)
  If AlreadyMovedTroopsThisPhase = True Then Exit Sub
  If TroopMoveSrc = 0 Then Exit Sub
  If Val(Land.TroopMoveNum.Text) = 0 Then Exit Sub
  If MyMap.Owner(CurrentMouse) <> Turn Then
    'This isn't our country.
    Land.Oopsie.Caption = "You must move troops to your own country."
    Call SNDBooBoo
    Exit Sub
  End If
  If CurrentMouse = TroopMoveSrc Then
    'We chose to move troops to the same country.  Goofy!
    Land.Oopsie.Caption = "You must move troops to a different country."
    Call SNDBooBoo
    Exit Sub
  End If
  If MyMap.TroopCount(CurrentMouse) >= MAX_COUNTRY_CAPACITY Then
    'This country is maxed out.
    Land.Oopsie.Caption = "This country is already maxed out.  The maximum allowable population of any country is 999."
    Call SNDBooBoo
    Exit Sub
  End If
  
  'We need to see if this country is adjacent to the last one.
  Done = False
  For i = 1 To MyMap.MaxNeighbors
    If MyMap.Neighbors(CurrentMouse, i) = TroopMoveSrc Then
      'We are a direct neighbor.
      Done = True
      Exit For
    ElseIf (MyMap.Neighbors(CurrentMouse, i) >= TILEVAL_COASTLINE) And ((MyMap.CountryType(TroopMoveSrc) And 1) = 1) Then
      'We need to check over water.
      For j = 1 To MyMap.MaxNeighbors
        If MyMap.Neighbors(CurrentMouse, i) = MyMap.Neighbors(TroopMoveSrc, j) Then
          'We can reach it over water.
          Done = True
          Exit For
        End If
      Next j
    End If
  If Done = True Then Exit For
  Next i
  
  If Done = False Then
    'We can't reach the country in question.
    Land.Oopsie.Caption = "Your troops cannot reach that country."
    Call SNDBooBoo
    Exit Sub
  End If
  
  If MyMap.TroopCount(CurrentMouse) + Val(Land.TroopMoveNum.Text) > MAX_COUNTRY_CAPACITY Then
    'We can't move all the troops in.  Only move as many as it takes to max it.
    Land.TroopMoveNum.Text = Trim(Str(MAX_COUNTRY_CAPACITY - MyMap.TroopCount(CurrentMouse)))
  End If
  
  'Now, we move troops from TroopMoveSrc to CurrentMouse!
  AlreadyMovedTroopsThisPhase = True
  'Encode the source country, dest country, and movement amount.
  Call SendClickToNetwork(Turn, Phase, (1000000 * Val(Land.TroopMoveNum.Text)) _
                         + (1000 * TroopMoveSrc) + CurrentMouse, 0)
  Call AddTroops(CurrentMouse, Val(Land.TroopMoveNum.Text))
  Call KillTroops(TroopMoveSrc, Val(Land.TroopMoveNum.Text))
  Call SNDTroopsIn(Player(Turn))
  Call ShortFlash(CurrentMouse)
  DrawMap
  Call NextTurn(Turn, Phase)
  Exit Sub
    
End Select

End Sub

Public Sub AnnexCountry(ThisCountry As Long, PlayerNumber As Integer)

'This sub claims a country for a player, makes it their color,
'and updates the player's totals.

Dim KillIt As Boolean

KillIt = False
MyMap.Owner(ThisCountry) = PlayerNumber
MyMap.CountryColor(ThisCountry) = Player(PlayerNumber)
NumOccupied(PlayerNumber) = NumOccupied(PlayerNumber) + 1
NumTroops(PlayerNumber) = NumTroops(PlayerNumber) + MyMap.TroopCount(ThisCountry)
'Let's make sure this is not an HQ by removing bit 2.
MyMap.CountryType(ThisCountry) = (MyMap.CountryType(ThisCountry) And 253)
'Let's destroy it's port if configured to do so.
Select Case CFGPorts
  Case 1:
    'Captured.  Do nothing.
    KillIt = False
  Case 2:
    'Destroyed.  Kill the port.
    KillIt = True
  Case 3:
    'Depends on how bad we beat up the other guy.
    If DefendStrength = 0 Then  'Don't want division by 0 below.
      KillIt = True
    ElseIf (AttackStrength / DefendStrength) > 4 Then
      KillIt = True
    End If
  Case 4:
    'Random.  If we are a network client, don't do this, wait for new data from the server.
    If MyNetworkRole <> NW_CLIENT Then
      If (Rnd(1) < 0.5) Then KillIt = True
    End If
End Select

'Kill the port if we're supposed to.
If KillIt = True Then MyMap.CountryType(ThisCountry) = (MyMap.CountryType(ThisCountry) And 254)

End Sub

Public Sub SecedeCountry(ThisCountry As Long)

'This sub updates a player's totals when a country is lost.

NumOccupied(MyMap.Owner(ThisCountry)) = NumOccupied(MyMap.Owner(ThisCountry)) - 1
NumTroops(MyMap.Owner(ThisCountry)) = NumTroops(MyMap.Owner(ThisCountry)) - MyMap.TroopCount(ThisCountry)

End Sub

Public Sub AddTroops(ToCountry As Long, Amount As Long)

'This sub just adds troops to a country.
MyMap.TroopCount(ToCountry) = MyMap.TroopCount(ToCountry) + Amount

'We need to make sure we didn't overshoot the limit!
If MyMap.TroopCount(ToCountry) > MAX_COUNTRY_CAPACITY Then MyMap.TroopCount(ToCountry) = MAX_COUNTRY_CAPACITY

End Sub

Public Sub KillTroops(FromCountry As Long, Amount As Long)

'This sub kills troops from a country.
MyMap.TroopCount(FromCountry) = MyMap.TroopCount(FromCountry) - Amount

'We need to make sure we didn't drop under zero!
If MyMap.TroopCount(FromCountry) < 0 Then MyMap.TroopCount(FromCountry) = 0

End Sub

Public Sub CalculateStrengths(AttackerTurn As Integer, DefenderCountry As Long)

Dim a As Long
Dim B As Long
Dim i As Integer
Dim j As Integer

    AttLndNum = 0
    AttWtrNum = 0
    DefLndNum = 0
    DefWtrNum = 0
    AttackStrength = 0
    DefendStrength = MyMap.TroopCount(DefenderCountry)
    WaterAttackStrength = 0
    WaterDefendStrength = 0
    
    For a = 1 To MyMap.NumberOfCountries
      If (MyMap.Owner(a) = AttackerTurn Or MyMap.Owner(a) = MyMap.Owner(DefenderCountry)) And a <> DefenderCountry Then
        'Let's only look at countries owned by one of the two combatants.
        For i = 1 To MyMap.MaxNeighbors
          'See if this country borders the contested country.
          If MyMap.Neighbors(a, i) <= 0 Then
            'We've checked all neighbors of this country.
            Exit For
          ElseIf (MyMap.Neighbors(a, i) = DefenderCountry) And (MyMap.TroopCount(a) > 0) Then
            'This country directly borders the contested country and
            'has troops in it which can attack or defend.
            If MyMap.Owner(a) = AttackerTurn And MyMap.Owner(DefenderCountry) <> AttackerTurn Then
              'This neighboring country belongs to the attacking player.
              AttackStrength = AttackStrength + MyMap.TroopCount(a)
              AttLndNum = AttLndNum + 1
            ElseIf MyMap.Owner(a) = AttackerTurn And MyMap.Owner(DefenderCountry) = AttackerTurn Then
              'This neighboring country belongs to the same person as the
              'country clicked on.  Probably a right-click.
              DefendStrength = DefendStrength + MyMap.TroopCount(a)
              DefLndNum = DefLndNum + 1
            ElseIf MyMap.Owner(a) = MyMap.Owner(DefenderCountry) Then
              'This neighboring country belongs to the defending player.
              DefendStrength = DefendStrength + MyMap.TroopCount(a)
              DefLndNum = DefLndNum + 1
            End If
            'Note that we exit the for loop here so that we don't count
            'bordering countries that share a water mass twice!
            Exit For
          ElseIf MyMap.Neighbors(a, i) >= TILEVAL_COASTLINE Then
            'Let's see if this country can reach the contested country by water.
            For j = 1 To MyMap.MaxNeighbors
              If MyMap.Neighbors(DefenderCountry, j) <= 0 Then
                'We've checked all the bodies of water we can get here by.
                Exit For
              'Let's see if we can make it there.  A port is necessary!
              ElseIf (MyMap.Neighbors(DefenderCountry, j) = MyMap.Neighbors(a, i)) _
                 And ((MyMap.CountryType(a) And 1) = 1) And (MyMap.TroopCount(a) > 0) Then
                'Country 'a' can reach the contested country by water!
                If MyMap.Owner(a) = AttackerTurn And MyMap.Owner(DefenderCountry) <> AttackerTurn Then
                  'The attacking player attacks over water.
                  WaterAttackStrength = WaterAttackStrength + MyMap.TroopCount(a)
                  AttWtrNum = AttWtrNum + 1
                ElseIf MyMap.Owner(a) = AttackerTurn And MyMap.Owner(DefenderCountry) = AttackerTurn Then
                  'We probably right-clicked.  Add to defense total for current player.
                  WaterDefendStrength = WaterDefendStrength + MyMap.TroopCount(a)
                  DefWtrNum = DefWtrNum + 1
                ElseIf MyMap.Owner(a) = MyMap.Owner(DefenderCountry) Then
                  'The defending player defends with boats.
                  WaterDefendStrength = WaterDefendStrength + MyMap.TroopCount(a)
                  DefWtrNum = DefWtrNum + 1
                End If
                'Exit the loops so we don't count two bodies of water.   :)
                i = MyMap.MaxNeighbors   'Fudging.
                Exit For
              End If
            Next j
          End If
        Next i
      End If
    Next a
    
'New for 2.0 BETA:  Add up all overseas countries and THEN multiply by ship modifier.
AttackStrength = AttackStrength + Int(WaterAttackStrength * ShipPct)
DefendStrength = DefendStrength + Int(WaterDefendStrength * ShipPct)
    
'A quick fudge to fix a defunct player with countries still out there.
'Although the player won't be able to right-click anymore, this will
'still affect computer AI.  If there's not a threat, don't react to it!
If NumOccupied(AttackerTurn) = -1 Then
  AttackStrength = 0
End If

End Sub

Public Sub ClaimHQ(ThatCountry As Long, PlayerNumber As Integer)

'This sub marks a country as the HQ for the current player.

Call AnnexCountry(ThatCountry, PlayerNumber)
MyMap.CountryType(ThatCountry) = MyMap.CountryType(ThatCountry) Or 2

End Sub

Public Sub Reinforce(ThatCountry As Long, Turn As Integer)

Dim TempStore As Long

TempStore = MyMap.TroopCount(ThatCountry)
Call AddTroops(ThatCountry, CFGBonusTroops * NumOccupied(Turn))
NumTroops(MyMap.Owner(ThatCountry)) = NumTroops(MyMap.Owner(ThatCountry)) + _
                    (MyMap.TroopCount(ThatCountry) - TempStore)

Call DrawMap

End Sub

Public Sub MakePort(ThatCountry As Long)

MyMap.CountryType(ThatCountry) = MyMap.CountryType(ThatCountry) Or 1
Call DrawMap

End Sub

Public Function Coastal(ThatCountry As Long) As Boolean

Dim i As Integer

Coastal = False
For i = 1 To MyMap.MaxNeighbors
  If MyMap.Neighbors(ThatCountry, i) >= TILEVAL_COASTLINE Then
    Coastal = True
    Exit For
  End If
Next i

End Function

Public Function AttackCountry(ThatCountry As Long, Turn As Integer)

Dim a As Long
Dim TempOwnerb As Integer
Dim TempNumber As Long

      'This sub assumes that the AttackStrength and DefendStrength have already
      'been calculated and that AttackStrength > DefendStrength.  The current
      'player is doing the attacking.
      
      'Update stats to show this country was attacked.
      StatScreen.UpdateAttackedStats Turn, MyMap.Owner(ThatCountry)
      
      'Let's kill some defenders!
      TempNumber = MyMap.TroopCount(ThatCountry)
      Call KillTroops(ThatCountry, AttackStrength - DefendStrength)
      'Let's update the player totals.
      If MyMap.TroopCount(ThatCountry) = 0 Then
        NumTroops(MyMap.Owner(ThatCountry)) = NumTroops(MyMap.Owner(ThatCountry)) - TempNumber
        StatScreen.UpdateKilledStats Turn, MyMap.Owner(ThatCountry), TempNumber  'Update stats.
      Else
        NumTroops(MyMap.Owner(ThatCountry)) = NumTroops(MyMap.Owner(ThatCountry)) - (AttackStrength - DefendStrength)
        StatScreen.UpdateKilledStats Turn, MyMap.Owner(ThatCountry), (AttackStrength - DefendStrength)  'Update stats.
      End If
      Call Land.Explode(ThatCountry, MyMap.CountryColor(ThatCountry))
      If (MyMap.TroopCount(ThatCountry) = 0) And ((MyMap.CountryType(ThatCountry) And 2) = 2) Then
        'We conquered an enemy HQ!
        'Update stats.
        TempOwnerb = MyMap.Owner(ThatCountry)
        StatScreen.UpdateDefeatedStats Turn, TempOwnerb
        StatScreen.UpdateOvertakenStats Turn, TempOwnerb
        NumOccupied(TempOwnerb) = -1
        NumTroops(TempOwnerb) = 0
        'This country gets claimed like normal.
        Call AnnexCountry(ThatCountry, Turn)
        'But we divvy up the remaining countries according to the config setting.
        For a = 1 To MyMap.NumberOfCountries
          If MyMap.Owner(a) = TempOwnerb Then
            'We need to do something with this country.
            Select Case CFGConquer
              Case 1:
                'The victor wins all taken countries.
                Call AnnexCountry(a, Turn)
                'Count this country as overtaken.
                StatScreen.UpdateOvertakenStats Turn, TempOwnerb
              Case 2:
                'All this player's countries turn neutral again.
                'Kill ports, leave troops.
                MyMap.Owner(a) = 0
                MyMap.CountryType(a) = 0
                'Assign it an 'unoccupied' color.
                MyMap.CountryColor(a) = CFGUnoccupiedColor
              Case 3:
                'Countries stay owned by the defunct player.
                'Basically, we do nothing here!  The remaining players
                'will need to attack and conquer the countries to claim them.
              Case 4:
                'All countries turn neutral and troops are destroyed.
                'Same as case 2 with troops reset.
                MyMap.Owner(a) = 0
                MyMap.CountryType(a) = 0
                MyMap.TroopCount(a) = 0
                'Assign it an 'unoccupied' color.
                MyMap.CountryColor(a) = CFGUnoccupiedColor
              Case 5:
                'Chaos erupts!
                'This is determined randomly.  Therefore, in a network game the server
                'will perform the random rolls and then all map data will be sent
                'to each client.  This is the best way to do this, since it's possible
                'for MANY countries to be affected.  We'll also have to update
                'overtaken statistics on the clients as well.
                If MyNetworkRole <> NW_CLIENT Then
                  'First, kill all ports and HQs.
                  MyMap.CountryType(a) = 0
                  'Now see if we kill troops in it.  40% chance.
                  If Rnd(1) < 0.4 Then
                    'Kill 'em all!
                    MyMap.TroopCount(a) = 0
                  End If
                  'There is a 50% chance of going neutral, a
                  '25% chance of victor claiming, and a
                  '25% chance of staying loyal.
                  If Rnd(1) < 0.5 Then
                    'Country goes neutral.
                    MyMap.Owner(a) = 0
                    MyMap.CountryColor(a) = CFGUnoccupiedColor
                  ElseIf Rnd(1) < 0.5 Then
                    'Victor claims it.
                    Call AnnexCountry(a, Turn)
                    'Count this country as overtaken.
                    StatScreen.UpdateOvertakenStats Turn, TempOwnerb
                  'Else it stays loyal, do nothing.
                  End If
                End If
              End Select
            End If
          Next a
        'Now send all map data to the clients if we are the server and a random
        'game parameter is in effect.
        If (MyNetworkRole = NW_SERVER) And (CFGConquer = 5) Then
          SendNewCountryData
          SendUpdatedStats
          Land.SetupIndicators
          Land.RedrawScreen
        End If
        Call SNDLargeExplosion
      ElseIf (MyMap.TroopCount(ThatCountry) = 0) And NumOccupied(MyMap.Owner(ThatCountry)) > 0 Then
        'We just conquered a player's country, but it wasn't the last one.
        StatScreen.UpdateOvertakenStats Turn, MyMap.Owner(ThatCountry)
        Call SecedeCountry(ThatCountry)
        Call AnnexCountry(ThatCountry, Turn)
        If (MyNetworkRole = NW_SERVER) And (CFGPorts = 4) Then
          SendPortStatus ThatCountry
          SendUpdatedStats
          Land.SetupIndicators
          DrawMap
        End If
        Call SNDMediumExplosion
      ElseIf (MyMap.TroopCount(ThatCountry) = 0) And NumOccupied(MyMap.Owner(ThatCountry)) = -1 Then
        'We're cleaning up countries owned by a now-defunct player.
        StatScreen.UpdateOvertakenStats Turn, MyMap.Owner(ThatCountry)
        Call AnnexCountry(ThatCountry, Turn)
        If (MyNetworkRole = NW_SERVER) And (CFGPorts = 4) Then
          SendPortStatus ThatCountry
          SendUpdatedStats
          Land.SetupIndicators
          DrawMap
        End If
        Call SNDMediumExplosion
      Else
        Call SNDSmallExplosion
      End If
      'If we conquered countries, we may need to remove old port graphics.
      Call Land.RedrawScreen

End Function
