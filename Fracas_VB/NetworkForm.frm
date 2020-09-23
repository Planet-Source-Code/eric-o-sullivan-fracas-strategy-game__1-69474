VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form NetworkForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Network Game"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3135
   Icon            =   "NetworkForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar NetProg 
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer HBtimer 
      Interval        =   1
      Left            =   2760
      Top             =   3000
   End
   Begin VB.CommandButton CancelBut 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton JoinBut 
      Caption         =   "Join!"
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox MyPlayerName 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox HostName 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   "BigSmozz"
      Top             =   2280
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock FracasSock 
      Index           =   0
      Left            =   2760
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label lblName 
      Caption         =   "Player Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblHost 
      Caption         =   "Hostname:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   11
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label ColorStatus 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label StatusText 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label ColorPool 
      Alignment       =   2  'Center
      Caption         =   "Color1"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "NetworkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Sub CancelBut_Click()

Dim i As Integer

On Error Resume Next

'Restore our PlayerType settings so that the Players dialog will be normal
'the next time this player goes there.
If MyNetworkRole = NW_CLIENT Then
  For i = 1 To MAX_PLAYERS
    PlayerType(i) = TempPlayerType(i)
  Next i
End If

'We can't join or start any games until this is done!
Land.MenuNew.Enabled = False
Land.MenuJoinGame.Enabled = False
Land.MenuSame.Enabled = False
Land.MenuLoad.Enabled = False
Land.MenuLoadSavedGame.Enabled = False
Land.MenuGetOptionsFromMapFile.Enabled = False
NetworkArbitrationInProgress = True

'Make us immediately invisible.
Me.Hide

'Get rid of the chat window.
Unload ChatForm

'Tell anyone who's listening that we're gone!
SendQuitToNetwork

'Give the Winsock controls time to send that message.
For i = 1 To 30
  DoEvents
  Sleep 20
Next i

NetworkArbitrationInProgress = False

'Make me a normal app, and get out of networking.
ActOnReturnCode NM_I_AM_QUITTING

End Sub

Private Sub Form_Load()

Dim i As Integer

InitNetwork

'Set up the network form based on our type.
If MyNetworkRole = NW_SERVER Then
  'We're the server.
  lblHost.Visible = False
  lblName.Visible = False
  HostName.Visible = False
  MyPlayerName.Visible = False
  JoinBut.Visible = False
  CancelBut.Visible = True
  
  SetUpColorList
  
  StatusText.Caption = "Waiting for players to join..."
  
  'Start listening for clients.
  NetworkState = NS_WAITING_FOR_CONNECTIONS
  ListenForClients
Else
  'We're a client.
  lblHost.Visible = True
  lblName.Visible = True
  HostName.Visible = True
  MyPlayerName.Visible = True
  JoinBut.Visible = True
  CancelBut.Visible = True
  NetworkState = NS_IDLE
  
  StatusText.Caption = "Enter server name or TCP/IP address to connect to and click Join."
  
  For i = 1 To 6
    ColorPool(i).Visible = False
    ColorStatus(i).Visible = False
  Next i

End If

NetProg.Visible = False

End Sub

Public Sub SetUpColorList()

Dim i As Integer

'Set up the colors.
For i = 1 To MAX_PLAYERS
  ColorPool(i).BackColor = PlayerColorCodes(Player(i))
  ColorPool(i).ForeColor = PlayerTextColor(Player(i))
  ColorPool(i).Caption = PlayerName(i)
  Select Case PlayerType(i)
    Case PTYPE_INACTIVE:
      ColorPool(i).Visible = False
      ColorStatus(i).Visible = False
    Case PTYPE_HUMAN:
      ColorPool(i).Visible = True
      ColorStatus(i).Visible = True
      If MyNetworkRole = NW_SERVER Then
        ColorStatus(i).Caption = "Local"
      Else
        ColorStatus(i).Caption = "Host"
      End If
    Case PTYPE_COMPUTER:
      ColorPool(i).Visible = True
      ColorStatus(i).Visible = True
      ColorStatus(i).Caption = "Computer"
    Case PTYPE_NETWORK:
      ColorPool(i).Visible = True
      ColorStatus(i).Visible = True
      If MyNetworkRole = NW_SERVER Then
        If MyNetIndex(i) = 0 Then
          ColorStatus(i).Caption = "Waiting..."
        Else
          ColorStatus(i).Caption = "Joined"
        End If
      Else
        ColorStatus(i).Caption = "Taken"
      End If
    Case PTYPE_NET_AVAIL:
      'Used to tell clients when a player slot is available.
      ColorPool(i).Visible = True
      ColorStatus(i).Visible = True
      ColorStatus(i).Caption = "Available"
  End Select
Next i

Form_Paint
DoEvents

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'Disable the close gadget for now.
If UnloadMode = 0 Then Cancel = 1

End Sub

Private Sub FracasSock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim PlayerId As Integer

'Just pass the request through to the network module.
ConnectRequest Index, requestID

End Sub

Private Sub FracasSock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim Pnum As Integer
Dim NetString As String
NetworkForm.FracasSock(Index).GetData NetString, vbString

If NetString = vbNullString Then Exit Sub

'We take the incoming string and simply append it to the end of
'the appropriate StringQ.
Pnum = Index
If MyNetworkRole = NW_CLIENT Then Pnum = 1

'Add it to the right queue.
ConcatenateStr Pnum, NetString

End Sub

Private Sub FracasSock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

MsgBox "A network error occurred:" & vbCrLf & vbCrLf & _
   Index & " - " & Number & " - " & Description
   
'Now we remove ourselves from the network.
CancelBut_Click

End Sub

Private Sub HBtimer_Timer()

'This is the entry point for the heartbeat timer.  This function is
'responsible for sending out the next message in each queue.  Important!

HeartBeat

End Sub

Private Sub JoinBut_Click()

'Tell the server that we're ready to go!
If (ConnectToHost(HostName.Text)) Then
  'Connection was accepted.
  NetworkState = NS_WAITING_FOR_JOIN_ACK
  StatusText.Caption = "Enter your name and click on an available color."
  JoinBut.Visible = False
  HostName.Visible = False
  lblHost.Visible = False
Else
  'Connection refused or bad host name, etc.
  StatusText.Caption = "Unable to connect.  Check Hostname and try again."
End If


End Sub

Private Sub ColorPool_Click(Index As Integer)

'The player clicked on one of the colors.  See if they're choosing.
'TODO:  Test that the Player name is valid.
If MyPlayerName.Text <> vbNullString Then
  PlayerChoseColor Index, MyPlayerName.Text
Else
  PlayerChoseColor Index, ColorPool(Index).Caption
End If

End Sub

Private Sub AdjustPlayerTypesForClient()

'This sub takes the current Player Types and adjusts them for this
'client.  Basically, this means making us a HUMAN and every other
'active player a SERVER controlled player.

Dim i As Integer

'Make every active player a SERVER player.  This means that we rely
'on the server to make this player's moves.  As a client, we don't care
'if the player is really a computer or another networked player --
'We wait on them regardless.
For i = 1 To MAX_PLAYERS
  If PlayerType(i) > PTYPE_INACTIVE Then
    PlayerType(i) = PTYPE_SERVER
  End If
Next i

'Mark US as a HUMAN.  This is so we can get input from this machine!
PlayerType(MyClientIndex) = PTYPE_HUMAN

End Sub

Public Sub ActOnReturnCode(ReturnCode As String)

'This sub performs actions on the network form based on messages we receive
'from other machines.

Dim i As Integer

If ReturnCode = vbNullString Then Exit Sub

Select Case ReturnCode
  Case NM_I_AM_QUITTING:
    'Make me a normal app, and get out of networking.
    For i = 0 To MAX_PLAYERS
      KillWinsockControl i
    Next i
    InitNetwork
    NetworkState = NS_IDLE
    MyNetworkRole = NW_NONE
    FracasSock(0).Close
    
    'TODO:  Set up the title screen if we aren't already there.
    Land.SetupMenusForGameOver
    MapMade = False
    GameMode = GM_BUILDING_TITLE

    Unload Me
    
  Case NM_YOU_ARE_GOOD_TO_GO:
    If MyNetworkRole = NW_SERVER Then
      'We just connected with all network players.  It's time to start.
      'Hide this form so our Winsock controls still work.
      Me.Hide
    ElseIf MyNetworkRole = NW_CLIENT Then
      'We just got our main config packet.  Need to leave so that our
      'map can be set up normally.  We'll be right back to get country
      'data though.
      Me.Hide
    End If
  Case NM_SETUP_COMPLETE:
    'Store off the PlayerType array, since it can change quite a bit depending on whether
    'we're a server or a client.  These will be restored when leaving a network game so
    'that we don't have any weirdness on the Players dialog.
    For i = 1 To MAX_PLAYERS
      TempPlayerType(i) = PlayerType(i)
    Next i
    If MyNetworkRole = NW_CLIENT Then
      'We've got all of our data and are ready to start!
      'We need to adjust player types here (this machine is human,
      'everyone else is controlled by the server) and then
      'return back to setup.
      AdjustPlayerTypesForClient
      Me.Hide
    Else
      'We're the server and everyone is here.  Just start.
      Me.Hide
    End If
End Select

End Sub

Public Sub UpdateProgressBar(Prog As Integer)

'This sub updates the progress bar we see during the transmission of map data.
'If Prog is 0, we're the server and we need to add up data for all clients.

Dim MapVert As Integer
Dim Progress As Single
Dim NumClients As Integer
Dim TotalLinesSent As Integer
Dim i As Integer

MapVert = MyMap.Ysize

If Prog = 0 Then
  'Count the clients and add up sent lines.
  NumClients = 0
  For i = 1 To MAX_PLAYERS
    If NetArray(i) > 0 Then
      NumClients = NumClients + 1
      TotalLinesSent = TotalLinesSent + DataArray(i)
    End If
  Next i
  'Progress is the percentage of everything sent to all clients.
  Progress = (TotalLinesSent / (MapVert * NumClients)) * 100
Else
  'Client.  Prog is how many lines of the map we've received.
  Progress = (Prog / MapVert) * 100
End If

NetProg.Value = Progress

End Sub

Private Sub Form_Paint()

Dim i As Integer
Const Xbase = 6
Const Ybase = 13
Const Dist = 24

NetworkForm.Cls

If NetworkState = NS_WAITING_FOR_USER_COLOR Then
  For i = 1 To MAX_PLAYERS
    If PlayerType(i) = PTYPE_NET_AVAIL Then
      u% = BitBlt(hdc, Xbase, Ybase + ((i - 1) * Dist), GFX_GRID, GFX_GRID, Land!LandMap.hdc, GFX_ICONS_X * GFX_GRID, (GFX_ICONS_Y + 1) * GFX_GRID, SRCAND)
      u% = BitBlt(hdc, Xbase, Ybase + ((i - 1) * Dist), GFX_GRID, GFX_GRID, Land!LandMap.hdc, GFX_ICONS_X * GFX_GRID, GFX_ICONS_Y * GFX_GRID, SRCINVERT)
    End If
  Next i
End If

End Sub
