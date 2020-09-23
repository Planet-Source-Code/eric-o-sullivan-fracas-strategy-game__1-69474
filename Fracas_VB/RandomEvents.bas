Attribute VB_Name = "RandomEvents"
Option Explicit
Option Base 1

Public Sub RollForRandomEvent(Turn As Integer)

Dim i As Long
Dim j As Long
Dim k As Long
Dim MyPct As Single

Dim EventNumber As Integer
Dim RandomCountryNumber As Long
Dim MessageNumber As Integer

'First, we calculate rankings.
'* The player in first place is not eligible for a random event.
'* The player in last place has a 2% chance of getting an event.
'If applicable:
'* All other players have a 1% chance.
'Thus, the better off you are, the less likely you are to have
'something good fall out of the sky!

'We don't roll for a random event if we're a client.  We get that information
'from the server.
If MyNetworkRole = NW_CLIENT Then Exit Sub

'No random events until after turn 5.  This allows the game to balance
'a bit before we skew it, hehehe.
If TurnCounter < 6 Then Exit Sub

'Get the current totals.
StatScreen.CalculateScores

'Calculate the Percentages for this player.
MyPct = (CFGEvents - 1) / 100
If CFGEvents = 3 Then MyPct = MyPct * 3  'Frequent means thrice as much!

'See if we're in first place!
EventNumber = 0
EventInProgress = False
If STATrank(Turn) = 1 Then
  'No events for you!  You're leading!
  EventInProgress = False
ElseIf STATrank(Turn) = NumPlayers Then
  'Twice the chance of random event.  You're in dead last!
  If Rnd(1) < (MyPct * 2) Then
    EventInProgress = True
    EventNumber = Int(Rnd(1) * 5) + 1    '1 thru 5.
  End If
Else
  'Normal chance of random event.  You're in the middle somewhere.
  If Rnd(1) < MyPct Then
    EventInProgress = True
    EventNumber = Int(Rnd(1) * 4) + 1    '1 thru 4.
  End If
End If

'DEBUG stuff
'EventInProgress = True
'EventNumber = Int(Rnd(1) * 5) + 1

If EventInProgress = False Then Exit Sub   'Leave if we don't have one this turn.

RandomCountryNumber = 0

'Now, see which event we're creating!
Select Case EventNumber
  
Case 0:
    Exit Sub   'No event for you!
  
Case 1:  'Gift of land.
  'We pick a totally random unoccupied plot of land and give it to this player.
  'This case just finds the country and stores its ID.
  For i = 0 To 150
    RandomCountryNumber = Int(Rnd(1) * MyMap.NumberOfCountries) + 1
    If MyMap.Owner(RandomCountryNumber) = 0 Then Exit For
  Next i
  If i > 150 Then
    'We couldn't find empty land -- so no special bonus.
    EventInProgress = False
    Exit Sub
  End If
  
Case 2:  'Coast Guard arrives.
  'We go through each country that this player owns and add a port if we can.
  'If we can add just one port, then we can do this event.
  Done = False
  For i = 1 To MyMap.NumberOfCountries
    If Done = True Then Exit For
    If MyMap.Owner(i) = Turn Then
      For j = 1 To MyMap.MaxNeighbors
        If MyMap.Neighbors(i, CInt(j)) >= TILEVAL_COASTLINE Then
          'We found a country that borders water.
          If ((MyMap.CountryType(i) And 1) <> 1) Then
            'And it doesn't yet have a port!  We can do this.
            Done = True
          End If
          Exit For
        End If
      Next j
    End If
  Next i
  If Done = False Then
    'Not bordering any water, or already have ports in the right spots.
    EventInProgress = False
    Exit Sub
  End If

Case 3:  '1st contact.
  'Pick a random country owned by this player and add CFGInitTroopCt times 25 to it.
  'Make sure we can find a country of theirs.
  For i = 0 To 150
    RandomCountryNumber = Int(Rnd(1) * MyMap.NumberOfCountries) + 1
    If (MyMap.Owner(RandomCountryNumber) = Turn) And (MyMap.TroopCount(RandomCountryNumber) + _
                                            (25 * CFGInitTroopCt) < MAX_COUNTRY_CAPACITY) Then Exit For
  Next i
  If i > 150 Then
    'We couldn't find a country where troops could fit -- so no special bonus.
    EventInProgress = False
    Exit Sub
  End If
  
Case 4:  'HQ bolstered.
  'Add a port plus (25 times CFGInitTroopCt) troops to this player's HQ.
  For i = 1 To MyMap.NumberOfCountries
    If (MyMap.Owner(i) = Turn) And ((MyMap.CountryType(i) And 2) = 2) Then
      'We found it.  First, add a port if not there.
      'Store the HQ in RandomCountryNumber so we don't have to find it again later.
      RandomCountryNumber = i
      Done = False
      'If there's no port or room for troops, then we'll do it.  Otherwise, no.
      If ((MyMap.CountryType(i) And 1) <> 1) Then Done = True
      If MyMap.TroopCount(i) + (25 * CFGInitTroopCt) < MAX_COUNTRY_CAPACITY Then Done = True
      If Done = False Then
        'Not enough space in the HQ and already had a port!
        EventInProgress = False
        Exit Sub
      End If
      Exit For
    End If
  Next i
  
Case 5:  'Population explosion!
  'We add 25% troops to all of this player's countries.
  Done = False
  For i = 1 To MyMap.NumberOfCountries
    If (MyMap.Owner(i) = Turn) And (MyMap.TroopCount(i) >= 4) Then
      If MyMap.TroopCount(i) + Int(0.25 * MyMap.TroopCount(i)) <= MAX_COUNTRY_CAPACITY Then
        'We found a country of theirs that can be added to.  We can do this event.
        Done = True
        Exit For
      End If
    End If
  Next i
  If Done = False Then
    'Already all maxed out!
    EventInProgress = False
    Exit Sub
  End If

End Select

MessageNumber = (Int(Rnd(1) * 3) + 1)

'At this point, we have the event number, which country it can occur on (if applicable),
'and which random message to throw at the player.

'So now send this event and event information to all clients!
If MyNetworkRole = NW_SERVER Then SendRandomEvent Turn, EventNumber, RandomCountryNumber, MessageNumber

'Finally, perform the event ourself.
ActivateRandomEvent Turn, EventNumber, RandomCountryNumber, MessageNumber
DrawMap
Land.UpdateMessages
Land.SetupIndicators
SNDBonusTwinkles

End Sub

Public Sub ActivateRandomEvent(Turn As Integer, EventNumber As Integer, RandomCountryNumber As Long, MessageNumber As Integer)

'This sub actually performs the passed random event.  This could come from the function above
'during a normal game, or from the SERVER in a networked game.

Dim i As Long
Dim j As Long

'Now, see which event we're creating!
Select Case EventNumber
  
Case 0:
    Exit Sub   'No event for you!
  
Case 1:  'Gift of land.
  'Give the passed random country to this player.
  Call AnnexCountry(RandomCountryNumber, Turn)
  Call BonusTwinkles(RandomCountryNumber, Player(Turn))
  'Random message.
  Select Case MessageNumber
  Case 1:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has inherited a plot of land from a rich uncle!"
  Case 2:
    Land.Oopsie.Caption = " The natives of " & MyMap.CountryName(RandomCountryNumber) & " have sworn allegiance to " & PlayerName(Turn) & "!"
  Case 3:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " is granted a bonus country this turn!"
  End Select
  
Case 2:  'Coast Guard arrives.
  'We go through each country that this player owns and add a port.
  For i = 1 To MyMap.NumberOfCountries
    If MyMap.Owner(i) = Turn Then
      For j = 1 To MyMap.MaxNeighbors
        If MyMap.Neighbors(i, CInt(j)) >= TILEVAL_COASTLINE Then
          'We found a country that borders water.
          If ((MyMap.CountryType(i) And 1) <> 1) Then
            'And it doesn't yet have a port!
            MyMap.CountryType(i) = MyMap.CountryType(i) Or 1
            Call SmallBonusTwinkles(i, Player(Turn))
          End If
          Exit For
        End If
      Next j
    End If
  Next i
  Select Case MessageNumber
  Case 1:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has instated a national coast guard!  All coastal countries gain ports."
  Case 2:
    Land.Oopsie.Caption = "Ports have been built on all available countries to protect the lands of " & PlayerName(Turn) & "!"
  Case 3:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has been granted naval power!  All coastal countries gain ports."
  End Select
  
Case 3:  '1st contact.
  'Add CFGInitTroopCt times 25 to the passed random country.
  Call AddTroops(RandomCountryNumber, (25 * CFGInitTroopCt))
  NumTroops(Turn) = NumTroops(Turn) + (25 * CFGInitTroopCt)
  Call BonusTwinkles(RandomCountryNumber, Player(Turn))
  Select Case MessageNumber
  Case 1:
  Land.Oopsie.Caption = " " & PlayerName(Turn) & " has made contact with an extraterrestrial race! " & _
                        Trim(Str(25 * CFGInitTroopCt)) & " colonists are staying behind."
  Case 2:
    Land.Oopsie.Caption = " " & Trim(Str(25 * CFGInitTroopCt)) & " alien refugees have been stranded here from another planet.  " & _
                        PlayerName(Turn) & " has put them to work!"
  Case 3:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has been 'visited' by an extraterrestrial race!"
  End Select
  
Case 4:  'HQ bolstered.
  'Add a port plus (25 times CFGInitTroopCt) troops to this player's HQ.
  'We stored the HQ's ID in RandomCountryNumber before we came here, remember?
  MyMap.CountryType(RandomCountryNumber) = MyMap.CountryType(RandomCountryNumber) Or 1
  Call AddTroops(RandomCountryNumber, 25 * CFGInitTroopCt)
  NumTroops(Turn) = NumTroops(Turn) + (25 * CFGInitTroopCt)
  Call BonusTwinkles(RandomCountryNumber, Player(Turn))
  Select Case MessageNumber
  Case 1:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has recruited peasants to bolster HQ defense."
  Case 2:
    Land.Oopsie.Caption = " A rich lord sympathetic to " & PlayerName(Turn) & " has offered to help defend HQ!"
  Case 3:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has been granted a boost to HQ!"
  End Select
  
Case 5:  'Population explosion!
  'We add 25% troops to all of this player's countries.
  Done = False
  For i = 1 To MyMap.NumberOfCountries
    If (MyMap.Owner(i) = Turn) And (MyMap.TroopCount(i) >= 4) Then
      If MyMap.TroopCount(i) + Int(0.25 * MyMap.TroopCount(i)) <= MAX_COUNTRY_CAPACITY Then
        MyMap.TroopCount(i) = MyMap.TroopCount(i) + Int(0.25 * MyMap.TroopCount(i))
        Call SmallBonusTwinkles(i, Player(Turn))
      End If
    End If
  Next i
  Land.CalculateTotals
  Select Case MessageNumber
  Case 1:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has distributed fertility pills to the unsuspecting populace!"
  Case 2:
    Land.Oopsie.Caption = " During other players' turns, the people of " & PlayerName(Turn) & " have been making babies!"
  Case 3:
    Land.Oopsie.Caption = " " & PlayerName(Turn) & " has experienced an unusual growth in population."
  End Select

End Select

End Sub

Public Sub RandomEventDoer()

'Basically, we just wait for all the little gems to bounce away.

If EventInProgress = False Then Exit Sub

If NoBalls = False Then Exit Sub

EventInProgress = False

End Sub

Public Function BonusTwinkles(CountryID As Long, ExpColor As Integer)

'Big bonus twinkle explosion for one country.
Call BuildBalls(BONUS_COUNT, _
                (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                BONUS_INTENSITY, _
                BONUS_SPREAD, _
                BONUS_ELASTIC, _
                BONUS_SIZE, _
                ExpColor)

End Function

Public Function SmallBonusTwinkles(CountryID As Long, ExpColor As Integer)

'Big bonus twinkle explosion for one country.
Call BuildBalls(SMBONUS_COUNT, _
                (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                SMBONUS_INTENSITY, _
                SMBONUS_SPREAD, _
                SMBONUS_ELASTIC, _
                SMBONUS_SIZE, _
                ExpColor)

End Function
