Attribute VB_Name = "NetworkPlay"
Option Explicit
Option Base 1

'Private NumberHeard As Integer

'What kind of network connection this instance of Fracas currently is.
Public MyNetworkRole As Integer
Public Const NW_NONE = 0
Public Const NW_SERVER = 1
Public Const NW_CLIENT = 2

Public NetworkState As Integer
Public Const NS_IDLE = 0
Public Const NS_WAITING_FOR_CONNECTIONS = 1  'Server
Public Const NS_WAITING_FOR_JOIN_ACK = 2     'Client
Public Const NS_WAITING_FOR_USER_COLOR = 3   'Client
Public Const NS_WAITING_TO_START = 4         'Client
Public Const NS_SENDING_SETUP_DATA = 5       'Server
Public Const NS_RECEIVING_SETUP_DATA = 6     'Client
Public Const NS_ACTIVE = 7                   'Server & Client

Public Const NM_HEARTBEAT = "BuhBump"
Public Const NM_ACK = "Roger"
Public Const NM_IDENTIFY_YOURSELF = "WhoAreYou"
Public Const NM_MY_IDENTITY = "HelloIAm"
Public Const NM_PLAYER_LIST = "PlayerList"
Public Const NM_I_CHOOSE_COLOR = "IWillBe"
Public Const NM_I_AM_BUILDING_THE_MAP = "ReticulatingSplines"
Public Const NM_YOU_ARE_GOOD_TO_GO = "HangOn"
Public Const NM_I_AM_QUITTING = "SeeYa"
Public Const NM_SETUP_DATA_1 = "OptionsAndStuff"
Public Const NM_SETUP_DATA_1_RCV = "GotFirstOne"
Public Const NM_SETUP_DATA_2 = "CountryData"
Public Const NM_SETUP_DATA_2_RCV = "GotCountryData"
Public Const NM_SETUP_DATA_3 = "LandNameData"
Public Const NM_SETUP_DATA_3_RCV = "GotLandNameData"
Public Const NM_SETUP_DATA_4 = "WaterNameData"
Public Const NM_SETUP_DATA_4_RCV = "GotWaterNameData"
Public Const NM_SETUP_DATA_MAP = "MapLine"
Public Const NM_SETUP_DATA_MAP_RCV = "GotMapLine"
Public Const NM_SETUP_COMPLETE = "StickAForkInMe"
Public Const NM_PLAYER_MOVE = "WeAreMoving"
Public Const NM_PLAYER_PASS = "NoThankYa"
Public Const NM_PLAYER_RESIGN = "Uncle"
Public Const NM_RENAME = "LetsCallThis"
Public Const NM_NEW_COUNTRY_DATA = "LotsOfStuffChanged"
Public Const NM_NEW_STATS = "StatUpdate"
Public Const NM_PORT_STATS = "NewPort"
Public Const NM_RANDOM_EVENT = "HangOnToYourHats"
Public Const NM_RETRY = "GimmeAnotherChance"    'Never sent to a partner.
Public Const NM_CHAT = "Pssst"

'NetArray contains the player number of each networked player.
'The index is the control index for that player.
'Used by servers only.
Public NetArray(MAX_PLAYERS) As Integer
Public DataArray(MAX_PLAYERS) As Integer
Public Const NA_PICKING_COLOR = 99   'Client.
Public NetworkArbitrationInProgress As Boolean   'Used to block any sort of message sending activity.
Public MyClientIndex As Integer  'Clients.
Public Const MAX_SEQ_NUM = 99999
Public Const MSG_Q_SIZE = 15  'Max messages that can be queued.
Public Const MSG_TIMEOUT = 250  'Number of iterations of HeartBeat to wait for an ACK.
Dim TurnInProgress As Boolean  'Used to prevent multiple turns over LAN.
Dim MsgQ(MAX_PLAYERS, MSG_Q_SIZE) As String
Dim StringQ(MAX_PLAYERS) As String
Dim Qtimer(MAX_PLAYERS) As Integer
Dim SeqNumber As Long   'Used to give each message a unique ID.

Public Sub ListenForClients()

'We're the server.
'The server is ready to start accepting connection requests.
If NetworkForm.FracasSock(0).State <> sckClosed Then InitNetwork
NetworkForm.FracasSock(0).LocalPort = 3737
NetworkForm.FracasSock(0).Listen

End Sub

Public Function ConnectToHost(HostName As String) As Boolean

ConnectToHost = False

'We're a client.
'Use control 0 to talk to the server and establish a connection.
If NetworkForm.FracasSock(0).State <> sckClosed Then
  InitNetwork
  NetworkForm.FracasSock(0).Close
  DoEvents
End If
NetworkForm.FracasSock(0).RemoteHost = HostName
NetworkForm.FracasSock(0).RemotePort = 3737
NetworkForm.FracasSock(0).Connect
Do
  DoEvents
Loop Until (NetworkForm.FracasSock(0).State = sckConnected) Or _
           (NetworkForm.FracasSock(0).State = sckError)

If NetworkForm.FracasSock(0).State = sckConnected Then ConnectToHost = True

End Function

Public Sub ConnectRequest(Index As Integer, requestID As Long)

' The server got a connection request from a client.  If we haven't
' yet heard from everyone, then we're still at the 'pick colors'
' dialog.  If we have heard from everyone already, then someone's
' trying to connect to a game in progress!  Reject it.

Dim NewConnectionNum As Integer
Dim i As Integer

NewConnectionNum = 0
If (NetworkState = NS_WAITING_FOR_CONNECTIONS) Then
  If Index = 0 Then
    'Server got a new connection.
    'Find the first available slot in our netarray.
    For i = 1 To MAX_PLAYERS
      If NetArray(i) = 0 Then
        NewConnectionNum = i
        Exit For
      End If
    Next i
    
    If (NewConnectionNum > 0) Then
      'We found a slot for this player to join.  Mark them as
      'Configuring their color.
      NetArray(NewConnectionNum) = NA_PICKING_COLOR
      Load NetworkForm.FracasSock(NewConnectionNum)
      NetworkForm.FracasSock(NewConnectionNum).LocalPort = 3737
      NetworkForm.FracasSock(NewConnectionNum).Accept requestID
      'Send this client the standard greeting. This is:
      '* Which Player Number we want him to be.
      '* The current list of players as it stands.
      'We will get back a response from this player indicating
      'that they agree.  At that point, we are bound to the client.
      SendMsg NewConnectionNum, NM_IDENTIFY_YOURSELF & "," & Trim(Str(NewConnectionNum))
      SendMsg NewConnectionNum, NM_PLAYER_LIST & "," & BuildPlayerList
    End If
  End If
End If

'If Index wasn't 0, then someone tried to connect to a client's
'Winsock control -- so do nothing.

End Sub

Private Function QuitRequest(NetIndex As Integer) As String

Dim i As Integer

QuitRequest = vbNullString  'Do nothing unless we decide otherwise.

If MyNetworkRole = NW_SERVER Then
  'A client just WUSSED out on us.  If we're still picking colors,
  'Set this player back to available.  If we're in the middle
  'of a game, let a computer take over.  If it's the end of a game,
  'It doesn't matter what we do here.
  If NetArray(NetIndex) > 0 Then
    'We got the message from someone with an active connection.
    'Close out their control.
    KillWinsockControl NetIndex
    'See if they're in a game.
    If GameMode = GM_GAME_ACTIVE Then
      'We're in the middle of a game.  This guy's a quitter!
      PlayerType(NetArray(NetIndex)) = PTYPE_COMPUTER
      Personality(NetArray(NetIndex)) = 1  'Stonewall for now.
      'Now clear their index.
      NetArray(NetIndex) = 0
    Else
      'We're picking colors.  Update our screen.
      'Clear their index first.
      NetArray(NetIndex) = 0
      NetworkForm.SetUpColorList
      'Tell everyone else that this player is a WUSS.
      BuildAndSendPlayerList
    End If
    DataArray(NetIndex) = 0
    For i = 1 To MSG_Q_SIZE
      MsgQ(NetIndex, i) = vbNullString
    Next i
  End If
Else
  'The server is leaving us.  We have no choice but to go away.
  If (GameMode = GM_GAME_ACTIVE) Or _
     (NetworkState = NS_WAITING_FOR_USER_COLOR) Or (NetworkState = NS_WAITING_TO_START) Then
    MsgBox "The server has left the network.  Fracas cannot continue as a client.", vbCritical, "No Server"
    QuitRequest = NM_I_AM_QUITTING
  Else
    'The server is probably leaving during the tribute dialog.  The game is already over.
  End If
End If

End Function

Public Function ParseIncomingMessage(NetIndex As Integer, NetString As String) As String

Dim i As Integer
Dim j As Long
Dim TempInt As Integer
Dim TempInt2 As Integer
Dim TempNum1 As Long
Dim TempNum2 As Long
Dim TempNum3 As Single
Dim TempNum4 As Long
Dim TempStr As String
Dim TempTurn As Integer
Dim TempPhase As Integer
Dim MySeqNum As Long
Dim MyCRC As Long

'This sub takes the passed message from the passed player and acts on it.
ParseIncomingMessage = vbNullString

'If the CRC check fails, we do NOTHING.
MyCRC = Val(ArgAt(NetString, 1))
NetString = StripOffFirstAt(NetString)
If MyCRC <> CalcCRC(NetString) Then Exit Function  'Bad CRC.

'Now, see if this is an ACK.
If Arg(NetString, 1) = NM_ACK Then
  'Yup, so let's get this msg off of our queue.
  MySeqNum = Val(Arg(NetString, 2))
  AckMsg NetIndex, MySeqNum
ElseIf Arg(NetString, 1) = NM_I_AM_QUITTING Then
  'We just got a quit request.
  ParseIncomingMessage = QuitRequest(NetIndex)
ElseIf ArgAt(NetString, 1) = NM_CHAT Then
  'It's a chat message.  Decode it and move on quick.
  'We just got a line of chat.
  TempInt = Val(ArgAt(NetString, 2))
  TempStr = ArgAt(NetString, 3)
  'Propagate it if we're the server!
  If MyNetworkRole = NW_SERVER Then SendChatToNetwork TempInt, TempStr, NetIndex
  ChatForm.ChatTextReceived TempInt, TempStr
Else
  'This isn't an ACK message, it's a real one!
  MySeqNum = Val(ArgAt(NetString, 1))
  NetString = StripOffFirstAt(NetString)
'--------------------------------------SERVER----------------------------------------
  If MyNetworkRole = NW_SERVER Then
    'We're the server and we just got a message from a client.
    Select Case Arg(NetString, 1)
      Case NM_HEARTBEAT:
        'Just a heartbeat.  Ack it.
        Acknowledge NetIndex, MySeqNum
      Case NM_I_CHOOSE_COLOR:
        'A player has chosen a color.  Check that it's valid, and
        'then tell them to wait.
        TempInt = Val(Arg(NetString, 2))
        If NetArray(NetIndex) = NA_PICKING_COLOR Then
          'This player is now this color.
          NetArray(NetIndex) = TempInt
          PlayerName(TempInt) = Arg(NetString, 3)
          'Tell this player that we accept.
          Acknowledge NetIndex, MySeqNum
          SendMsg NetIndex, NM_YOU_ARE_GOOD_TO_GO & "," & Trim(Str(TempInt))
          'Update everyone else of this fact.
          BuildAndSendPlayerList
          'Update our own screen.
          NetworkForm.SetUpColorList
          'Check to see if we have everyone.  If so, start the game!
          TempInt = 0
          For i = 1 To MAX_PLAYERS
            If PlayerType(i) = PTYPE_NETWORK And _
                 ((MyNetIndex(i) = 0) Or (MyNetIndex(i) = NA_PICKING_COLOR)) Then
              'Found a network player that hasn't joined yet or is in
              'the process of picking their color.
              TempInt = 1
              Exit For
            End If
          Next i
          If TempInt = 0 Then
            'We have everyone.
            NetworkForm.StatusText.Caption = "Sending game data.  Please wait."
            'We don't want people to cancel while sending data.  That's a misuse
            'case that could get really, really messy.
            NetworkForm.CancelBut.Visible = False
            'Send all players a heads-up that we're building the map now.
            For i = 1 To MAX_PLAYERS
              If NetArray(i) > 0 Then
                SendMsg i, NM_I_AM_BUILDING_THE_MAP
              End If
            Next i
            ParseIncomingMessage = NM_YOU_ARE_GOOD_TO_GO
          End If
        Else
          'Tell the client that this isn't acceptable.  On the other
          'hand, we might just do nothing here, since the client's list
          'will be updated soon.  Basically, two players probably
          'tried to pick the same color at the same time.
          Acknowledge NetIndex, MySeqNum
        End If
      Case NM_SETUP_DATA_1_RCV:
        'This client got the first round of setup data.
        Acknowledge NetIndex, MySeqNum
        SendMsg NetIndex, NM_SETUP_DATA_2 & "," & BuildCountryData
      Case NM_SETUP_DATA_2_RCV:
        'This client got the second round of setup data.
        Acknowledge NetIndex, MySeqNum
        SendMsg NetIndex, NM_SETUP_DATA_3 & "," & BuildLandNameData
      Case NM_SETUP_DATA_3_RCV:
        'This client got the third round of setup data.
        Acknowledge NetIndex, MySeqNum
        SendMsg NetIndex, NM_SETUP_DATA_4 & "," & BuildWaterNameData
      Case NM_SETUP_DATA_4_RCV:
        'This client got their final data packet.  Now send them the first line
        'of map info.  We will only send a line of map data when the previous
        'line was confirmed.  If we don't confirm a line, we'll rely on our
        'heartbeat timer to detect loss of comms and resend.
        Acknowledge NetIndex, MySeqNum
        SendMsg NetIndex, BuildMapLine(1)
      Case NM_SETUP_DATA_MAP_RCV:
        'This client just told us they received the last line of map we sent.
        Acknowledge NetIndex, MySeqNum
        TempInt = Val(Arg(NetString, 2))
        'Increment our sent counter.  When everyone's sent counter is equal
        'to the Y size of the map, we can start!
        DataArray(NetIndex) = TempInt
        'Update our progress bar.
        NetworkForm.UpdateProgressBar (0)
        'Go to the next line.
        TempInt = TempInt + 1
        If TempInt > MyMap.Ysize Then
          'We're done with this guy!  See if anyone else is still waiting on data.
          SendMsg NetIndex, NM_SETUP_COMPLETE
          TempNum1 = 0
          For i = 1 To MAX_PLAYERS
            If (NetArray(i) > 0) And (DataArray(i) < MyMap.Ysize) Then
              TempNum1 = 1
              Exit For
            End If
          Next i
          If TempNum1 = 0 Then
            'Everyone has all of their map data!  Time to kick off the game.
            NetworkState = NS_ACTIVE
            ParseIncomingMessage = NM_SETUP_COMPLETE   'Return code to start the game.
          End If
        Else
          'Send out the next line.
          SendMsg NetIndex, BuildMapLine(TempInt)
        End If
      Case NM_SETUP_COMPLETE:
        'This client has confirmed that we gave them everything they need.
        Acknowledge NetIndex, MySeqNum
        SendMsg NetIndex, NM_HEARTBEAT
      Case NM_PLAYER_MOVE:
        'We just got a move request from a client.
        ParseIncomingMessage = NM_RETRY  'We retry by default unless we succeed.
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        TempNum1 = Val(Arg(NetString, 4))  'Country num.
        'Leave if it isn't this player's turn (probably a retried msg)
        If Land.GetTurn <> TempTurn Then Exit Function
        'Now just pretend that this click really happened.
        If NumOccupied(TempTurn) = 0 Then
          'This player is choosing their country.
          Acknowledge NetIndex, MySeqNum
          SendClickToNetwork TempTurn, TempPhase, TempNum1, NetIndex
          ChooseHQProc TempNum1, TempTurn
          Call NextTurn(TempTurn, TempPhase)
          ParseIncomingMessage = vbNullString
          Exit Function
        End If
        Select Case TempPhase
          Case 1:
            'We got a reinforce request.  Make sure we're in the reinforce phase.
            If Land.GetPhase <> TempPhase Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the reinforce.
            SendClickToNetwork TempTurn, TempPhase, TempNum1, NetIndex
            ReinforceProc TempNum1, TempTurn
            Call NextPhase(TempTurn, TempPhase)
          Case 2:
            'This is an action request.  Make sure we're in the action phase.
            If Land.GetPhase <> TempPhase Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the action.
            SendClickToNetwork TempTurn, TempPhase, TempNum1, NetIndex
            ActionProc TempNum1, TempTurn
            Call NextPhase(TempTurn, TempPhase)
          Case 3:
            'We should *never* get here.  No netmsgs sent during phase 3.
            MsgBox "An error occurred during troop movement source select." & vbCrLf & _
                   "Please Email jmerlo@austin.rr.com and report this bug!"
            End
          Case 4:
            'This is a troop movement request.  Make sure we're in the right phase.
            If Land.GetPhase <> 3 Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the troop movement.
            TempNum3 = Int(TempNum1 / 1000000)
            TempNum4 = TempNum1 - (TempNum3 * 1000000)
            TempNum2 = Int(TempNum4 / 1000)
            TempNum4 = TempNum4 - (TempNum2 * 1000)
            SendClickToNetwork TempTurn, TempPhase, TempNum1, NetIndex
            TroopMoveProc TempNum2, TempNum4, TempNum3, TempTurn
            Call NextTurn(TempTurn, TempPhase)
          End Select
          ParseIncomingMessage = vbNullString
          TurnInProgress = False
      Case NM_PLAYER_PASS:
        'We just got a pass request from a client.
        ParseIncomingMessage = NM_RETRY  'We retry by default unless we succeed.
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        'Leave if it isn't this player's turn (probably a retried msg)
        If (Land.GetTurn <> TempTurn) Then Exit Function
        'Now do the pass.
        If TempPhase < 3 Then
          If Land.GetPhase <> TempPhase Then Exit Function
          TurnInProgress = True
          Acknowledge NetIndex, MySeqNum
          SendPassToNetwork TempTurn, TempPhase, NetIndex
          Call SNDWhistle
          Call NextPhase(TempTurn, TempPhase)
        Else
          'If we're somewhere in troop movement, just skip it all.
          If Land.GetPhase <> 3 Then Exit Function
          TurnInProgress = True
          Acknowledge NetIndex, MySeqNum
          SendPassToNetwork TempTurn, TempPhase, NetIndex
          Call SNDWhistle
          Call NextTurn(TempTurn, TempPhase)
        End If
        ParseIncomingMessage = vbNullString
        TurnInProgress = False
      Case NM_PLAYER_RESIGN:
        'We just got a resign request from a client.  That weenie!
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        'Leave if it isn't this player's turn (probably a retried msg)
        If (Land.GetTurn <> TempTurn) Then Exit Function
        If (Land.GetPhase <> TempPhase) Then Exit Function
        'Now do the resign.
        TurnInProgress = True
        Acknowledge NetIndex, MySeqNum
        SendResignToNetwork TempTurn, TempPhase, NetIndex
        Land.ResignProc TempTurn
        Call NextTurn(TempTurn, TempPhase)
        TurnInProgress = False
      Case NM_RENAME:
        'We just got a request to rename an entity, land or water.
        'This one is easy, just grab the ID and name and make it so.
        Acknowledge NetIndex, MySeqNum
        TempNum1 = Val(Arg(NetString, 2))   'Entity ID.
        TempStr = Arg(NetString, 3)    'New name for this entity.
        If TempNum1 < TILEVAL_COASTLINE Then
          'We're renaming land.
          MyMap.CountryName(TempNum1) = TempStr
        Else
          'We're renaming water.
          MyMap.WaterName(TempNum1 - 1000) = TempStr
        End If
        SendRenameToNetwork TempNum1, TempStr, NetIndex
        
    End Select
'--------------------------------------CLIENT----------------------------------------
  ElseIf MyNetworkRole = NW_CLIENT Then
    'We're a client and the server just sent us a message.
    Select Case Arg(NetString, 1)
      Case NM_HEARTBEAT:
        'Just a heartbeat.  Ack it.
        Acknowledge NetIndex, MySeqNum
      Case NM_IDENTIFY_YOURSELF:
        'Identify myself.  This will just put us into the
        'proper state for picking colors.
        NetworkState = NS_WAITING_FOR_USER_COLOR
        Acknowledge NetIndex, MySeqNum
      Case NM_PLAYER_LIST:
        'Update the player list on our dialog.
        For i = 1 To 6
          PlayerName(i) = Arg(NetString, ((i - 1) * 3) + 2)
          Player(i) = Val(Arg(NetString, ((i - 1) * 3) + 3))
          PlayerType(i) = Val(Arg(NetString, ((i - 1) * 3) + 4))
        Next i
        NetworkForm.SetUpColorList
        Acknowledge NetIndex, MySeqNum
      Case NM_YOU_ARE_GOOD_TO_GO:
        'The server has accepted our request to be this player.
        'When we get the go signal, MyIndex will become our
        'player number.
        Acknowledge NetIndex, MySeqNum
        MyClientIndex = Val(Arg(NetString, 2))
        SNDPlayFanfare Player(MyClientIndex)
        NetworkState = NS_WAITING_TO_START
        NetworkForm.lblName.Visible = False
        NetworkForm.MyPlayerName.Visible = False
        NetworkForm.StatusText.Caption = "Please wait for others to join."
      Case NM_I_AM_BUILDING_THE_MAP:
        'The server has everyone joined and is now building the map.
        'We'll just put up a message to that effect and disallow cancelling.
        Acknowledge NetIndex, MySeqNum
        NetworkForm.CancelBut.Visible = False
        NetworkForm.StatusText.Caption = "The server is building the map." & vbCrLf & "Please wait."
      Case NM_SETUP_DATA_1:
        'We just got our first setup packet.
        If NetworkState = NS_WAITING_TO_START Then
          NetworkForm.StatusText.Caption = "Receiving data.  Please wait."
          NetworkState = NS_RECEIVING_SETUP_DATA
          CFGCountrySize = Val(Arg(NetString, 2))
          CFGLandPct = Val(Arg(NetString, 3))
          CFGIslands = Val(Arg(NetString, 4))
          CFGLakeSize = Val(Arg(NetString, 5))
          CFGProportion = Val(Arg(NetString, 6))
          CFGShape = Val(Arg(NetString, 7))
          CFGBorders = Val(Arg(NetString, 8))
          CFGInitTroopPl = Val(Arg(NetString, 9))
          CFGInitTroopCt = Val(Arg(NetString, 10))
          CFGBonusTroops = Val(Arg(NetString, 11))
          CFGShips = Val(Arg(NetString, 12))
          CFGPorts = Val(Arg(NetString, 13))
          CFGConquer = Val(Arg(NetString, 14))
          CFGAISpeed = Val(Arg(NetString, 15))
          CFG1st = Val(Arg(NetString, 16))
          CFGHQSelect = Val(Arg(NetString, 17))
          CFGUnoccupiedColor = Val(Arg(NetString, 18))
          CFGResolution = Val(Arg(NetString, 19))
          CFGConquer = Val(Arg(NetString, 20))
          CFGCountrySize = Val(Arg(NetString, 21))
          TempNum1 = Val(Arg(NetString, 22))   'Number of countries.
          TempNum2 = Val(Arg(NetString, 23))   'LakeCode (num of water masses).\
          Land.SetTurn Val(Arg(NetString, 24))
          'Put our progress bar up...
          NetworkForm.NetProg.Value = 0
          NetworkForm.NetProg.Visible = True
          'Update our menus to reflect the saved config settings.
          Call ClearMenuChecks
          Call UpdateMenus
          'Now that we have a new resolution, see if it's valid.
          If CFGResolution > MaxAllowedResolution Then
            'Can't display a map at that resolution.  Exit gracefully.
            MsgBox "The game you are joining is using a map too big for your desktop.", vbOKOnly Or vbCritical, "Map Too Large"
            'Send back a code that says we can't go on.  Also signal
            'the server that we're leaving.
            SendMsg 0, NM_I_AM_QUITTING
            ParseIncomingMessage = NM_I_AM_QUITTING
            Exit Function
          Else
            'Resolution is valid, set up map.
            Land.ChangeRes
            MyMap.RedimensionStuff TempNum1, TempNum2
          End If
          'Signal the server that we got this much...
          Acknowledge NetIndex, MySeqNum
          SendMsg 0, NM_SETUP_DATA_1_RCV
        End If
      Case NM_SETUP_DATA_2:
        If NetworkState = NS_RECEIVING_SETUP_DATA Then
          Acknowledge NetIndex, MySeqNum
          'Parse country data.
          For j = 1 To MyMap.NumberOfCountries
            TempStr = Arg(NetString, j + 1)
            MyMap.CountryColor(j) = Val(ArgAt(TempStr, 1))
            MyMap.CountryType(j) = Val(ArgAt(TempStr, 2))
            MyMap.Owner(j) = Val(ArgAt(TempStr, 3))
            MyMap.TroopCount(j) = Val(ArgAt(TempStr, 4))
          Next j
          SendMsg 0, NM_SETUP_DATA_2_RCV
        End If
      Case NM_SETUP_DATA_3:
        If NetworkState = NS_RECEIVING_SETUP_DATA Then
          Acknowledge NetIndex, MySeqNum
          'Parse land names.
          For j = 1 To MyMap.NumberOfCountries
            MyMap.CountryName(j) = Arg(NetString, j + 1)
          Next j
          SendMsg 0, NM_SETUP_DATA_3_RCV
        End If
      Case NM_SETUP_DATA_4:
        If NetworkState = NS_RECEIVING_SETUP_DATA Then
          Acknowledge NetIndex, MySeqNum
          'Parse water names.
          For i = 1001 To MyMap.LakeCode
            MyMap.WaterName(i - 1000) = Arg(NetString, (i - 1000) + 1)
          Next i
          SendMsg 0, NM_SETUP_DATA_4_RCV
        End If
      Case NM_SETUP_DATA_MAP:
        If NetworkState = NS_RECEIVING_SETUP_DATA Then
          Acknowledge NetIndex, MySeqNum
          'The server just sent us a line of the map.  Put it in our array and
          'ask for the next.
          TempInt = Val(Arg(NetString, 2))   'This is the line number.
          For i = 1 To MyMap.Xsize
            MyMap.Grid(i, TempInt) = Val(Arg(NetString, i + 2))
          Next i
          'Update our progress bar.
          NetworkForm.UpdateProgressBar (TempInt)
          'Now tell the server we got this line.
          SendMsg 0, NM_SETUP_DATA_MAP_RCV & "," & Trim(Str(TempInt))
        End If
      Case NM_SETUP_COMPLETE:
        If NetworkState = NS_RECEIVING_SETUP_DATA Then
          Acknowledge NetIndex, MySeqNum
          'The server says we're finished!
          SendMsg 0, NM_SETUP_COMPLETE
          'Some map parameters need to be calculated.
          Call MyMap.FinishUpLoad
          'Recalculate ShipPct.  Fudge.
          Call Land.CollectMenuSettings(True)
          ParseIncomingMessage = NM_SETUP_COMPLETE   'Return code to start the game.
          NetworkState = NS_ACTIVE
        End If
      Case NM_PLAYER_MOVE:
        'We just got a move request from the server.
        ParseIncomingMessage = NM_RETRY  'We retry by default unless we succeed.
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        TempNum1 = Val(Arg(NetString, 4))  'Country num.
        'Special for clients:  Ack this message if it's the server
        'telling US where WE should move.  Obviously came from the
        'SendClickToNetwork function...
        If TempTurn = MyClientIndex Then
          Acknowledge NetIndex, MySeqNum
          ParseIncomingMessage = vbNullString
          Exit Function
        End If
        'Leave if it isn't this player's turn (probably a retried msg)
        If (Land.GetTurn <> TempTurn) Then Exit Function
        'Now just pretend that this click really happened.
        If NumOccupied(TempTurn) = 0 Then
          'This player is choosing their country.
          Acknowledge NetIndex, MySeqNum
          ChooseHQProc TempNum1, TempTurn
          Call NextTurn(TempTurn, TempPhase)
          ParseIncomingMessage = vbNullString
          Exit Function
        End If
        Select Case TempPhase
          Case 1:
            'We got a reinforce request.  Make sure we're in the reinforce phase.
            If Land.GetPhase <> TempPhase Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the reinforce.
            ReinforceProc TempNum1, TempTurn
            Call NextPhase(TempTurn, TempPhase)
          Case 2:
            'This is an action request.  Make sure we're in the action phase.
            If Land.GetPhase <> TempPhase Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the action.
            ActionProc TempNum1, TempTurn
            Call NextPhase(TempTurn, TempPhase)
          Case 3:
            'We should *never* get here.  No netmsgs set during phase 3.
            MsgBox "An error occurred during troop movement source select." & vbCrLf & _
                   "Please Email jmerlo@austin.rr.com and report this bug!"
            End
          Case 4:
            'This is a troop movement request.  Make sure we're in the right phase.
            If Land.GetPhase <> 3 Then Exit Function
            TurnInProgress = True
            Acknowledge NetIndex, MySeqNum
            'Now do the troop movement.
            TempNum3 = Int(TempNum1 / 1000000)
            TempNum1 = TempNum1 - (TempNum3 * 1000000)
            TempNum2 = Int(TempNum1 / 1000)
            TempNum1 = TempNum1 - (TempNum2 * 1000)
            TroopMoveProc TempNum2, TempNum1, TempNum3, TempTurn
            Call NextTurn(TempTurn, TempPhase)
        End Select
        ParseIncomingMessage = vbNullString
        TurnInProgress = False
      Case NM_PLAYER_PASS:
        'We just got a pass request from the server.
        ParseIncomingMessage = NM_RETRY  'We retry by default unless we succeed.
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        'Special for clients:  Ack this message if it's the server
        'telling US that WE should pass.  Obviously came from the
        'SendPassToNetwork function...
        If TempTurn = MyClientIndex Then
          Acknowledge NetIndex, MySeqNum
          ParseIncomingMessage = vbNullString
          Exit Function
        End If
        'Leave if it isn't this player's turn (probably a retried msg)
        If (Land.GetTurn <> TempTurn) Then Exit Function
        'Now do the pass.
        If TempPhase < 3 Then
          If Land.GetPhase <> TempPhase Then Exit Function
          TurnInProgress = True
          Acknowledge NetIndex, MySeqNum
          Call SNDWhistle
          Call NextPhase(TempTurn, TempPhase)
        Else
          'If we're somewhere in troop movement, just skip it all.
          If Land.GetPhase <> 3 Then Exit Function
          TurnInProgress = True
          Acknowledge NetIndex, MySeqNum
          Call SNDWhistle
          Call NextTurn(TempTurn, TempPhase)
        End If
        ParseIncomingMessage = vbNullString
        TurnInProgress = False
      Case NM_PLAYER_RESIGN:
        'We just got a resign request from the server.
        If TurnInProgress = True Then Exit Function
        TempTurn = Val(Arg(NetString, 2))
        TempPhase = Val(Arg(NetString, 3))
        'Special for clients:  Ack this message if it's the server
        'telling US that WE should resign.  Obviously came from the
        'SendPassToNetwork function...
        If TempTurn = MyClientIndex Then
          Acknowledge NetIndex, MySeqNum
          Exit Function
        End If
        'Leave if it isn't this player's turn (probably a retried msg)
        If (Land.GetTurn <> TempTurn) Then Exit Function
        If (Land.GetPhase <> TempPhase) Then Exit Function
        'Now do the resign.
        TurnInProgress = True
        Acknowledge NetIndex, MySeqNum
        Land.ResignProc TempTurn
        Call NextTurn(TempTurn, TempPhase)
        TurnInProgress = False
      Case NM_RENAME:
        'We just got a request to rename an entity, land or water.
        'This one is easy, just grab the ID and name and make it so.
        Acknowledge NetIndex, MySeqNum
        TempNum1 = Val(Arg(NetString, 2))   'Entity ID.
        TempStr = Arg(NetString, 3)    'New name for this entity.
        If TempNum1 < TILEVAL_COASTLINE Then
          'We're renaming land.
          MyMap.CountryName(TempNum1) = TempStr
        Else
          'We're renaming water.
          MyMap.WaterName(TempNum1 - 1000) = TempStr
        End If
      Case NM_NEW_COUNTRY_DATA:
        'Parse country data.  Similar to setup data 2.
        Acknowledge NetIndex, MySeqNum
        For j = 1 To MyMap.NumberOfCountries
          TempStr = Arg(NetString, j + 1)
          MyMap.CountryColor(j) = Val(ArgAt(TempStr, 1))
          MyMap.CountryType(j) = Val(ArgAt(TempStr, 2))
          MyMap.Owner(j) = Val(ArgAt(TempStr, 3))
          MyMap.TroopCount(j) = Val(ArgAt(TempStr, 4))
        Next j
        Land.CalculateTotals
        Land.SetupIndicators
        Land.RedrawScreen
      Case NM_NEW_STATS:
        'Just parse out the statistics.
        Acknowledge NetIndex, MySeqNum
        For i = 1 To MAX_PLAYERS
          TempStr = Arg(NetString, (i * 3) - 1)
          For TempInt = 1 To MAX_PLAYERS
            STATattacked(i, TempInt) = Val(ArgAt(TempStr, TempInt))
          Next TempInt
          TempStr = Arg(NetString, (i * 3))
          For TempInt = 1 To MAX_PLAYERS
            STATovertaken(i, TempInt) = Val(ArgAt(TempStr, TempInt))
          Next TempInt
          TempStr = Arg(NetString, (i * 3) + 1)
          For TempInt = 1 To MAX_PLAYERS
            STATkilled(i, TempInt) = Val(ArgAt(TempStr, TempInt))
          Next TempInt
        Next i
      Case NM_PORT_STATS:
        'Update a single country's port status.  Used for random port destruction.
        Acknowledge NetIndex, MySeqNum
        TempNum1 = Val(Arg(NetString, 2))
        TempInt = Val(Arg(NetString, 3))
        MyMap.CountryType(TempNum1) = TempInt
        Land.RedrawScreen
      Case NM_RANDOM_EVENT:
        TempTurn = Val(Arg(NetString, 2))
        If (TempTurn <> Land.GetTurn) Then Exit Function
        Acknowledge NetIndex, MySeqNum
        TempInt = Val(Arg(NetString, 3))  'Event Number.
        TempNum1 = Val(Arg(NetString, 4)) 'Random Country Number.
        TempInt2 = Val(Arg(NetString, 5)) 'Message Number.
        'Finally, perform the event ourself.
        ActivateRandomEvent TempTurn, TempInt, TempNum1, TempInt2
        DrawMap
        Land.UpdateMessages
        Land.SetupIndicators
        SNDBonusTwinkles
        
    End Select
    
  End If

End If

End Function

Public Sub InitNetwork()

'This sub just initializes everything.

Dim i As Integer
Dim j As Integer

NetworkArbitrationInProgress = False
TurnInProgress = False
For i = 1 To MAX_PLAYERS
  For j = 1 To MSG_Q_SIZE
    MsgQ(i, j) = vbNullString
  Next j
  StringQ(i) = vbNullString
  Qtimer(i) = 0
  NetArray(i) = 0
  DataArray(i) = 0
Next i

End Sub

Private Sub BuildAndSendPlayerList()

'Server.  This function builds and sends the string to tell all
'clients what the player distribution looks like.

Dim i As Integer
Dim TempStr As String

TempStr = BuildPlayerList

For i = 1 To MAX_PLAYERS
  If NetArray(i) > 0 Then
    SendMsg i, NM_PLAYER_LIST & "," & TempStr
  End If
Next i

End Sub

Private Function BuildPlayerList() As String

Dim TempStr As String
Dim i As Integer

TempStr = vbNullString

For i = 1 To MAX_PLAYERS
  'Each player has a name, a color, and a type.
  TempStr = TempStr & PlayerName(i) & ","
  TempStr = TempStr & Trim(Str(Player(i))) & ","
  If PlayerType(i) = PTYPE_NETWORK Then
    If MyNetIndex(i) = 0 Then
      'This slot hasn't been filled yet.
      TempStr = TempStr & Trim(Str(PTYPE_NET_AVAIL))
    Else
      'Someone's already snagged this color.
      TempStr = TempStr & Trim(Str(PTYPE_NETWORK))
    End If
  Else
    TempStr = TempStr & Trim(Str(PlayerType(i)))
  End If
  If i < 6 Then TempStr = TempStr & ","
Next i

BuildPlayerList = TempStr

End Function

Public Function MyNetIndex(PlayerNum As Integer)

'This function searches NetArray for the passed player number.
'If it's in the array, we return the network index of that player.

Dim i As Integer

MyNetIndex = 0
For i = 1 To MAX_PLAYERS
  If (NetArray(i) = PlayerNum) Then
    MyNetIndex = i
    Exit For
  End If
Next i

End Function

Public Function PlayerChoseColor(Index As Integer, MyName As String)

'The player clicked on one of the colors in the list.  If we are
'a client, in the picking stage, and that color is available,
'then that's who we'll be!  Send the server a message saying so.

If (MyNetworkRole = NW_CLIENT) And (NetworkState = NS_WAITING_FOR_USER_COLOR) Then
  'We are a client in the right state.  See if this color is valid.
  If PlayerType(Index) = PTYPE_NET_AVAIL Then
    'This is us!  Send the server a message, including our name.
    SendMsg 0, NM_I_CHOOSE_COLOR & "," & Trim(Str(Index)) & "," & MyName
  Else
    'Error message in status line.
    SNDBooBoo
    NetworkForm.StatusText.Caption = "That color is unavailable.  Choose another."
  End If
End If

End Function

Public Function KillWinsockControl(Index As Integer)

'This sub closes the specified Winsock connection and unloads it if
'it's not index 0 (0 is always present).

On Error Resume Next

'Get out if we're trying to mess with a control that doesn't exist for a client.
If MyNetworkRole = NW_CLIENT And Index > 0 Then Exit Function

If NetworkForm.FracasSock(Index).State = sckConnected Then
  'Close the connection.
  NetworkForm.FracasSock(Index).Close
  DoEvents
End If

'Unload this client's connection to us.  (We're obviously the server)
If Index > 0 Then Unload NetworkForm.FracasSock(Index)

End Function

Public Sub SendStartupDataToAll()

'Server.  Send each of our clients all of the necessary data to start
'up a game.

Dim TempStr As String
Dim i As Integer

'Change our state...
NetworkState = NS_SENDING_SETUP_DATA

'Put our progress bar up...
NetworkForm.NetProg.Value = 0
NetworkForm.NetProg.Visible = True

'Send every bit of pertinent setup data.  But not preference data.
'Also include the map dimensions so clients can initialize properly.
TempStr = Trim(Str(CFGCountrySize)) & "," & _
          Trim(Str(CFGLandPct)) & "," & _
          Trim(Str(CFGIslands)) & "," & _
          Trim(Str(CFGLakeSize)) & "," & _
          Trim(Str(CFGProportion)) & "," & _
          Trim(Str(CFGShape)) & "," & _
          Trim(Str(CFGBorders)) & "," & _
          Trim(Str(CFGInitTroopPl)) & "," & _
          Trim(Str(CFGInitTroopCt)) & "," & _
          Trim(Str(CFGBonusTroops)) & "," & _
          Trim(Str(CFGShips)) & "," & _
          Trim(Str(CFGPorts)) & "," & _
          Trim(Str(CFGConquer)) & "," & _
          Trim(Str(CFGAISpeed)) & "," & _
          Trim(Str(CFG1st)) & "," & _
          Trim(Str(CFGHQSelect)) & "," & _
          Trim(Str(CFGUnoccupiedColor)) & "," & _
          Trim(Str(CFGResolution)) & "," & _
          Trim(Str(CFGConquer)) & "," & _
          Trim(Str(CFGCountrySize)) & "," & _
          Trim(Str(MyMap.NumberOfCountries)) & "," & _
          Trim(Str(MyMap.LakeCode)) & "," & _
          Trim(Str(Land.GetTurn))
          
For i = 1 To MAX_PLAYERS
  If NetArray(i) > 0 Then
    'This control is in use.  Send them the startup data.
    SendMsg i, NM_SETUP_DATA_1 & "," & TempStr
  End If
Next i

End Sub

Private Function BuildMapLine(LineNum As Integer) As String

'Server.  This function packs one line of map data into a string for
'transmission to a client.

Dim TempStr As String
Dim i As Integer
Dim j As Integer

'The first argument is the line number...
TempStr = NM_SETUP_DATA_MAP & "," & Trim(Str(LineNum)) & ","
'Then each piece of map data.
For i = 1 To MyMap.Xsize
  TempStr = TempStr & Trim(Str(MyMap.Grid(i, LineNum)))
  If i < MyMap.Xsize Then TempStr = TempStr & ","
Next i

BuildMapLine = TempStr

End Function

Private Function BuildCountryData() As String

'This function puts all country data for this map into a string.

Dim i As Long
Dim TempStr As String

TempStr = vbNullString
For i = 1 To MyMap.NumberOfCountries
  TempStr = TempStr & Trim(Str(MyMap.CountryColor(i))) & "@" & _
                      Trim(Str(MyMap.CountryType(i))) & "@" & _
                      Trim(Str(MyMap.Owner(i))) & "@" & _
                      Trim(Str(MyMap.TroopCount(i)))
  If i < MyMap.NumberOfCountries Then TempStr = TempStr & ","
Next i

BuildCountryData = TempStr

End Function

Private Function BuildLandNameData() As String

'This function puts all land name information for this map into a string.

Dim i As Long
Dim TempStr As String

'Land names.
TempStr = vbNullString
For i = 1 To MyMap.NumberOfCountries
  TempStr = TempStr & MyMap.CountryName(i)
  If i < MyMap.NumberOfCountries Then TempStr = TempStr & ","
Next i

BuildLandNameData = TempStr

End Function

Private Function BuildWaterNameData() As String

'This function puts all water name information for this map into a string.

Dim i As Long
Dim TempStr As String

'Water names.
TempStr = vbNullString
For i = 1001 To MyMap.LakeCode
  TempStr = TempStr & MyMap.WaterName(i - 1000)
  If i < MyMap.LakeCode Then TempStr = TempStr & ","
Next i

BuildWaterNameData = TempStr

End Function

Public Sub SendClickToNetwork(Turn As Integer, Phase As Integer, CountryNum As Long, NotThisPlayer As Integer)

'This sub determines if we need to send this click to any of our clients
'or to the server.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We need to send this click to all clients.
  'Either a human at the server just moved, or we are propagating a client's
  'move to the other clients.
  For i = 1 To MAX_PLAYERS
    If (NetArray(i) > 0) And (i <> NotThisPlayer) Then
      SendMsg i, NM_PLAYER_MOVE & "," & Trim(Str(Turn)) & "," & _
                 Trim(Str(Phase)) & "," & Trim(Str(CountryNum))
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Send this click to the server for distribution to everyone.
  SendMsg 0, NM_PLAYER_MOVE & "," & Trim(Str(Turn)) & "," & _
             Trim(Str(Phase)) & "," & Trim(Str(CountryNum))
End If

End Sub

Public Sub SendPassToNetwork(Turn As Integer, Phase As Integer, NotThisPlayer As Integer)

'This sub determines if we need to send this pass to any of our clients
'or to the server.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We need to send this pass to all clients.
  'Either a human at the server just passed, or we are propagating a client's
  'pass to the other clients.
  For i = 1 To MAX_PLAYERS
    If (NetArray(i) > 0) And (i <> NotThisPlayer) Then
      SendMsg i, NM_PLAYER_PASS & "," & Trim(Str(Turn)) & "," & _
                 Trim(Str(Phase))
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Send this pass to the server for distribution to everyone.
  SendMsg 0, NM_PLAYER_PASS & "," & Trim(Str(Turn)) & "," & _
             Trim(Str(Phase))
End If

End Sub

Public Sub SendResignToNetwork(Turn As Integer, Phase As Integer, NotThisPlayer As Integer)

'This sub determines if we need to send this resign to any of our clients
'or to the server.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We need to send this resign to all clients.
  'Either a human at the server just resigned, or we are propagating a client's
  'resign to the other clients.
  For i = 1 To MAX_PLAYERS
    If (NetArray(i) > 0) And (i <> NotThisPlayer) Then
      SendMsg i, NM_PLAYER_RESIGN & "," & Trim(Str(Turn)) & "," & _
                 Trim(Str(Phase))
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Send this resign to the server for distribution to everyone.
  SendMsg 0, NM_PLAYER_RESIGN & "," & Trim(Str(Turn)) & "," & _
             Trim(Str(Phase))
End If

End Sub

Public Sub SendRenameToNetwork(Entity As Long, NewName As String, NotThisPlayer As Integer)

'This sub determines if we need to send this rename request to any of our clients
'or to the server.  Note that this sub handles renames of both land and water.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We need to send this rename to all clients.
  'Either a human at the server just renamed something, or we are propagating a client's
  'rename request to the other clients.
  For i = 1 To MAX_PLAYERS
    If (NetArray(i) > 0) And (i <> NotThisPlayer) Then
      SendMsg i, NM_RENAME & "," & Trim(Str(Entity)) & "," & NewName
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Send this rename request to the server for distribution to everyone.
  SendMsg 0, NM_RENAME & "," & Trim(Str(Entity)) & "," & NewName
End If

End Sub

Public Sub SendQuitToNetwork()

'This sub gets run when someone cancels the network form.
'Send out the quit message if anyone is listening.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We're the server.  Tell all clients that we're gone.
  For i = 1 To MAX_PLAYERS
    If NetArray(i) > 0 Then
      If NetworkForm.FracasSock(i).State = sckConnected Then
        NetworkForm.FracasSock(i).SendData "$" & AddCRC(NM_I_AM_QUITTING) & "%"
        DoEvents
      End If
    End If
  Next i
Else
  'We're a client.  Tell the server we're quitting.
  If NetworkForm.FracasSock(0).State = sckConnected Then
    NetworkForm.FracasSock(0).SendData "$" & AddCRC(NM_I_AM_QUITTING) & "%"
    DoEvents
  End If
End If

End Sub

Public Sub SendChatToNetwork(PlayerNum As Integer, ChatText As String, NotThisPlayer As Integer)

'This sub determines if we need to send this chat line to any of our clients
'or to the server.

Dim i As Integer

If MyNetworkRole = NW_SERVER Then
  'We need to send this line of chat to all clients.
  For i = 1 To MAX_PLAYERS
    If (NetArray(i) > 0) And (i <> NotThisPlayer) Then
      SendChatStr i, PlayerNum, ChatText
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Send this line of chat to the server for distribution to everyone.
  SendChatStr 0, PlayerNum, ChatText
End If

End Sub

Private Sub SendChatStr(Index As Integer, PlayerNum As Integer, ChatText As String)

Dim ChatStr As String

ChatStr = vbNullString

If NetworkForm.FracasSock(Index).State = sckConnected Then
  '@ signs instead of commas so we can send commas in our chats!
  ChatStr = NM_CHAT & "@" & Trim(Str(PlayerNum)) & "@" & ChatText
  ChatStr = AddCRC(ChatStr)
  NetworkForm.FracasSock(Index).SendData "$" & ChatStr & "%"
  DoEvents
End If

End Sub

Private Sub SendMsg(Index As Integer, Msg As String)

'This sub adds the passed message to the appropriate message queue.
'We rely on the net timer to send the first message in the queue, and
'messages are only removed from the queue when an ACK is received.

Dim Pnum As Integer
Dim i As Integer
Dim Done As Boolean

If Msg = vbNullString Then Exit Sub

'First, get a new sequence number and append it to the start of the message.
'This will be stripped off by the recipient.
SeqNumber = SeqNumber + 1
If SeqNumber > MAX_SEQ_NUM Then SeqNumber = 1
Msg = Trim(Str(SeqNumber)) & "@" & Msg

'Then calculate a checksum for this message and append it to the front.
Msg = AddCRC(Msg)

'Now add this message to the passed player's queue.
Pnum = Index
If MyNetworkRole = NW_CLIENT Then Pnum = 1
Done = False

'Loop forever and ever until we can put this message in the queue!
'Hopefully, we only go through this loop once.  But if a queue is full to the
'brim, we *CAN'T* drop the message.  So we'll freeze ourselves until the
'recipient acknowledges something in the queue and frees up space.
Do
  For i = 1 To MSG_Q_SIZE
    If MsgQ(Pnum, i) = vbNullString Then
      'We found the first empty slot, so put the message there.
      MsgQ(Pnum, i) = Msg
      Done = True
      'If this is the first message we're sending, make sure it goes out quick.
      If i = 1 Then Qtimer(Pnum) = 0
      Exit For
    End If
  Next i
  If Done = False Then DoEvents   'Give our NetTimer time to process sends.
Loop Until Done = True

End Sub

Private Sub AckMsg(Index As Integer, SeqNum As Long)

'This sub removes the message with the given sequence number from the
'passed player's message queue.  ONLY THE FIRST message in a queue can
'be acked!  If this message is never acked, we have a PROBLEM.

Dim Pnum As Integer
Dim i As Integer
Dim TempStri As String

TempStri = vbNullString
Pnum = Index
If MyNetworkRole = NW_CLIENT Then Pnum = 1

'We check the second @ argument for the seq num because the first is the CRC!
If ArgAt(MsgQ(Pnum, 1), 2) = Trim(Str(SeqNum)) Then
  'Get this message text.
  TempStri = ArgAt(MsgQ(Pnum, 1), 3)
  'Let's bump the whole queue up until this message text is no longer first.
  'This will ACK all messages in the queue that are retries of the same message.
  Do
    For i = 2 To MSG_Q_SIZE
      MsgQ(Pnum, i - 1) = MsgQ(Pnum, i)
    Next i
    MsgQ(Pnum, MSG_Q_SIZE) = vbNullString
  Loop Until ArgAt(MsgQ(Pnum, 1), 3) <> TempStri
  'Also reset our queue timer to immediately send the next one!
  Qtimer(Pnum) = 0
Else
  'Tried to ACK another message besides the first.  Bad, bad, bad!
  MsgBox "An Acknowledgement error occurred.  Please Email jmerlo@austin.rr.com and report this bug!", vbCritical, "ACK error"
  End
End If

End Sub

Public Sub HeartBeat()

On Error Resume Next

Dim i As Integer
Dim ReturnCode As String
Dim NetString As String

'First, we see if our StringQ has a complete message in it.
'If so, strip it off the front.
If MyNetworkRole = NW_SERVER Then
  'We need to check every one of our Queues.
  For i = 1 To MAX_PLAYERS
    NetString = StringQMaintenance(i)
    If NetString <> vbNullString Then
      'Now do whatever this message is telling us to do.  If we get a RETRY return code,
      'it means that we are lagging behind our partner and should keep trying until
      'our turn and phase catch up.
      Do
        ReturnCode = ParseIncomingMessage(i, NetString)
        If ReturnCode = NM_RETRY Then DoEvents
      Loop Until ReturnCode <> NM_RETRY
      NetworkForm.ActOnReturnCode (ReturnCode)
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Our queue is number 1.
  NetString = StringQMaintenance(1)
  If NetString <> vbNullString Then
    'Retry until we can successfully parse this message.
    Do
      ReturnCode = ParseIncomingMessage(i, NetString)
      If ReturnCode = NM_RETRY Then DoEvents
    Loop Until ReturnCode <> NM_RETRY
    NetworkForm.ActOnReturnCode (ReturnCode)
  End If
End If

'This is where we send the next message in each queue we hold.
If MyNetworkRole = NW_SERVER Then
  For i = 1 To MAX_PLAYERS
    If NetArray(i) > 0 Then
      'Check that the Winsock connection is good before sending.
      If NetworkForm.FracasSock(i).State = sckConnected Then
        Qtimer(i) = Qtimer(i) - 1
        If Qtimer(i) < 1 Then
          Qtimer(i) = MSG_TIMEOUT
          If MsgQ(i, 1) <> vbNullString Then
            NetworkForm.FracasSock(i).SendData "$" & MsgQ(i, 1) & "%"
            DoEvents
          End If
        End If
      End If
    End If
  Next i
ElseIf MyNetworkRole = NW_CLIENT Then
  'Check that the Winsock connection is good before sending.
  If NetworkForm.FracasSock(0).State = sckConnected Then
    Qtimer(1) = Qtimer(1) - 1
    If Qtimer(1) < 1 Then
      Qtimer(1) = MSG_TIMEOUT
      If MsgQ(i, 1) <> vbNullString Then
        NetworkForm.FracasSock(0).SendData "$" & MsgQ(1, 1) & "%"
        DoEvents
      End If
    End If
  End If
End If

'Debug purposes.  TODO:  Jason remove this!
UpdateDebugForm

End Sub

Private Function StringQMaintenance(Index As Integer) As String

'This function strips the first complete message from the front of the queue.
'Complete messages start with a dollar sign ($) and end with a percent (%).

Dim MyStr As String
Dim Pos As Integer

'Get out if no string.
StringQMaintenance = vbNullString
MyStr = StringQ(Index)
If MyStr = vbNullString Then Exit Function

'Find where the first $ is.
Pos = InStr(MyStr, "$")
If Pos = 0 Then
  'No dollar sign at all.  That's bad.  Clean this string out.
  StringQ(Index) = vbNullString
  Exit Function
ElseIf Pos > 1 Then
  'The first dollar sign is *not* at the front of the queue.
  'We need to chop off the leading stuff and start there.
  StringQ(Index) = Right(MyStr, Len(MyStr) - Pos + 1)
  MyStr = StringQ(Index)
End If

'If we got here then the first dollar sign is at the front of the queue.
'This is normal.  Now find if we have an end cap in there somewhere.
Pos = InStr(MyStr, "%")
If Pos = 0 Then
  'No end cap.  The rest of this line will be coming soon.
  Exit Function
End If

'If we got here then we have a complete message between 1 and Pos.
StringQMaintenance = Mid(MyStr, 2, Pos - 2)
StringQ(Index) = Right(MyStr, Len(MyStr) - Pos)

End Function

Private Sub Acknowledge(Index As Integer, SeqNum As Long)

'This function sends an ACK out to a networked machine.

Dim Pnum As Integer
Dim AckStr As String

AckStr = vbNullString
Pnum = Index
If MyNetworkRole = NW_CLIENT Then Pnum = 0

If NetworkForm.FracasSock(Pnum).State = sckConnected Then
  AckStr = NM_ACK & "," & Trim(Str(SeqNum))
  AckStr = AddCRC(AckStr)
  NetworkForm.FracasSock(Pnum).SendData "$" & AckStr & "%"
  DoEvents
End If

End Sub

Private Function StripOffFirstAt(MyStr As String) As String

'This sub strips the leading sequence number off of the passed string.
'There is an AT sign (@) between the sequence number and the rest of the
'message.

Dim i As Integer
Dim Pos As Integer

'Default return value is the passed string itself.
StripOffFirstAt = MyStr

Pos = InStr(MyStr, "@")
If Pos > 0 Then
  StripOffFirstAt = Right(MyStr, Len(MyStr) - Pos)
End If

End Function

Private Function AddCRC(MyStr As String) As String

'This function calculates a CRC for the passed string and prepends it to
'the same string, which then gets returned.

Dim CRC As Long
Dim i As Integer

CRC = CalcCRC(MyStr)
AddCRC = Trim(Str(CRC)) & "@" & MyStr

End Function

Private Function CalcCRC(MyStr As String) As Long

Dim CRC As Long
Dim i As Integer

CRC = 0
For i = 1 To Len(MyStr)
  CRC = CRC + Asc(Mid(MyStr, i, 1))
Next i

CalcCRC = CRC

End Function

Public Sub ConcatenateStr(Index As Integer, Frag As String)

'This sub just sticks the passed fragment on the end of the indicated queue.
StringQ(Index) = StringQ(Index) & Frag

End Sub

Public Sub SendNewCountryData()

'Server.  This sub sends out all country data to all clients.  Used after a calamatous
'random event like killing an HQ with the 'chaos erupts' option.
Dim i As Integer
Dim TempStr As String

TempStr = vbNullString
TempStr = BuildCountryData

For i = 1 To MAX_PLAYERS
  If (NetArray(i) > 0) Then
    SendMsg i, NM_NEW_COUNTRY_DATA & "," & TempStr
  End If
Next i

End Sub

Public Sub SendPortStatus(ThisCountry As Long)

'Server.  This sub updates a single country's port status.  Used for random port destruction.
Dim i As Integer

For i = 1 To MAX_PLAYERS
  If (NetArray(i) > 0) Then
    SendMsg i, NM_PORT_STATS & "," & Trim(Str(ThisCountry)) & "," & Trim(Str(MyMap.CountryType(ThisCountry)))
  End If
Next i

End Sub

Public Sub SendUpdatedStats()

'Server.  This sub updates the overtaken stats on the stat screen after someone's HQ
'goes down with the 'chaos erupts' option chosen.
Dim i As Integer
Dim TempStrg As String

TempStrg = vbNullString
TempStrg = BuildStats

For i = 1 To MAX_PLAYERS
  If (NetArray(i) > 0) Then
    SendMsg i, NM_NEW_STATS & "," & TempStrg
  End If
Next i

End Sub

Private Function BuildStats() As String

'Package up our overtaken stats.
Dim i As Integer
Dim TempStr As String

TempStr = vbNullString
For i = 1 To MAX_PLAYERS
  TempStr = TempStr & STATattacked(i, 1) & "@" & STATattacked(i, 2) & "@" & STATattacked(i, 3) & "@" & _
                      STATattacked(i, 4) & "@" & STATattacked(i, 5) & "@" & STATattacked(i, 6) & ","
  TempStr = TempStr & STATovertaken(i, 1) & "@" & STATovertaken(i, 2) & "@" & STATovertaken(i, 3) & "@" & _
                      STATovertaken(i, 4) & "@" & STATovertaken(i, 5) & "@" & STATovertaken(i, 6) & ","
  TempStr = TempStr & STATkilled(i, 1) & "@" & STATkilled(i, 2) & "@" & STATkilled(i, 3) & "@" & _
                      STATkilled(i, 4) & "@" & STATkilled(i, 5) & "@" & STATkilled(i, 6)
  If i < MAX_PLAYERS Then TempStr = TempStr & ","
Next i

BuildStats = TempStr

End Function

Public Sub SendRandomEvent(Turn As Integer, EventNumber As Integer, RandomCountryNumber As Long, MessageNumber As Integer)

'Server.  Sends random event information down to all clients for immedate processing
'before a person's turn can get too far underway.

Dim i As Integer

For i = 1 To MAX_PLAYERS
  If (NetArray(i) > 0) Then
    SendMsg i, NM_RANDOM_EVENT & "," & Trim(Str(Turn)) & "," & Trim(Str(EventNumber)) & "," & _
                                       Trim(Str(RandomCountryNumber)) & "," & Trim(Str(MessageNumber))
  End If
Next i

End Sub

Private Sub UpdateDebugForm()

'This DEBUG sub just puts all our relevant data on the debug form.
'TODO: Jason remove this!

If DebugForm.Visible = False Then Exit Sub

Dim i As Integer
Dim j As Integer

DebugForm.XMyClientIndex = MyClientIndex
DebugForm.XSeqNumber = SeqNumber
DebugForm.XTurnInProgress = TurnInProgress
DebugForm.XTurn = Land.GetTurn
DebugForm.XPhase = Land.GetPhase

For i = 1 To MAX_PLAYERS
  DebugForm.XNetArray(i) = NetArray(i)
  DebugForm.XDataArray(i) = DataArray(i)
  DebugForm.XQtimer(i) = Qtimer(i)
  If StringQ(i) <> vbNullString Then DebugForm.XStringQ(i) = StringQ(i)
Next i

For j = 1 To MSG_Q_SIZE
  If j = 1 Then
    If MsgQ(1, j) <> vbNullString Then DebugForm.XMsg1(j) = MsgQ(1, j)
    If MsgQ(2, j) <> vbNullString Then DebugForm.XMsg2(j) = MsgQ(2, j)
    If MsgQ(3, j) <> vbNullString Then DebugForm.XMsg3(j) = MsgQ(3, j)
    If MsgQ(4, j) <> vbNullString Then DebugForm.XMsg4(j) = MsgQ(4, j)
    If MsgQ(5, j) <> vbNullString Then DebugForm.XMsg5(j) = MsgQ(5, j)
    If MsgQ(6, j) <> vbNullString Then DebugForm.XMsg6(j) = MsgQ(6, j)
  Else
    DebugForm.XMsg1(j) = MsgQ(1, j)
    DebugForm.XMsg2(j) = MsgQ(2, j)
    DebugForm.XMsg3(j) = MsgQ(3, j)
    DebugForm.XMsg4(j) = MsgQ(4, j)
    DebugForm.XMsg5(j) = MsgQ(5, j)
    DebugForm.XMsg6(j) = MsgQ(6, j)
  End If
Next j

End Sub
