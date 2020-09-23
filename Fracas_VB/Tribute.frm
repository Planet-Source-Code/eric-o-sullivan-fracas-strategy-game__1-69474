VERSION 5.00
Begin VB.Form Tribute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4470
   Icon            =   "Tribute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton HiScoreButton 
      Caption         =   "View Hi Scores"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton StatsButton 
      Caption         =   "View Statistics"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton WinningButton 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label WinningSentence 
      Alignment       =   2  'Center
      Caption         =   "Sentencing"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label WinningPlayer 
      Alignment       =   2  'Center
      Caption         =   "Winning Player!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Tribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Form_Load()

Dim WinnerPl As Integer
Dim i As Integer

'Somebody won!  Let's give them a salute.

DoEvents    'Let the form come up clean...

If MyNetworkRole <> NW_NONE Then
  'Turn off high scores if part of a network game.
  HiScoreButton.Enabled = False
  'Also send out the final country configuration and stats if we're the server.
  If MyNetworkRole = NW_SERVER Then
    SendNewCountryData
    SendUpdatedStats
  End If
Else
  HiScoreButton.Enabled = True
End If

Call SNDApplause

'Let's get the exact turn of the person who won.
For i = 1 To 6
  If PlayerType(i) > PTYPE_INACTIVE And NumOccupied(i) > -1 Then WinnerPl = i
Next i

DoEvents

'Set up the big text.
WinningPlayer.Caption = PlayerName(WinnerPl) & " wins!"
WinningPlayer.BackColor = PlayerColorCodes(Player(WinnerPl))
WinningPlayer.ForeColor = PlayerTextColor(Player(WinnerPl))

'Set up the sentencing.
Select Case Int(Rnd(1) * 8) + 1
  Case 1:
  WinningSentence.Caption = "The people of " & PlayerName(WinnerPl) & " have chosen to spare the lives of " & _
                            "their enemies.  However, they will rot in prison for the next millenium!"
  Case 2:
  WinningSentence.Caption = PlayerName(WinnerPl) & " reigns supreme!  The rest of you upstarts had better " & _
                            "get back to your farmwork!"
  Case 3:
  WinningSentence.Caption = "As a gift on this glorious day of celebration, " & PlayerName(WinnerPl) & " has " & _
                            "decided not to behead the rest of you."
  Case 4:
  WinningSentence.Caption = "The enemies of " & PlayerName(WinnerPl) & " have been found guilty of ... well ... " & _
                            "something ... and will be hung at sunup."
  Case 5:
  WinningSentence.Caption = "The enemies of " & PlayerName(WinnerPl) & " will be made to pay for their crimes " & _
                            "against the new empire.  Long live " & PlayerName(WinnerPl) & "!"
  Case 6:
  WinningSentence.Caption = PlayerName(WinnerPl) & " has conquered the world!  " & _
                            "Well -- this little part of it, at least..."
  Case 7:
  WinningSentence.Caption = PlayerName(WinnerPl) & " has achieved demigod status by crushing the opposing " & _
                            "armies into submission."
  Case 8:
  WinningSentence.Caption = PlayerName(WinnerPl) & " has implemented a dictatorship in the newly conquered lands.  " & _
                            "All bow before the new king!"
End Select

'Set up the humbling button.
Select Case Int(Rnd(1) * 8) + 1
  Case 1:
  WinningButton.Caption = "Long live " & PlayerName(WinnerPl) & "!"
  Case 2:
  WinningButton.Caption = PlayerName(WinnerPl) & " reigns supreme!"
  Case 3:
  WinningButton.Caption = PlayerName(WinnerPl) & " is our master!"
  Case 4:
  WinningButton.Caption = "Until next time..."
  Case 5:
  WinningButton.Caption = "Click here to play again."
  Case 6:
  WinningButton.Caption = "Three cheers for " & PlayerName(WinnerPl) & "!"
  Case 7:
  WinningButton.Caption = "Sleep with one eye open, " & PlayerName(WinnerPl) & "..."
  Case 8:
  WinningButton.Caption = "All hail " & PlayerName(WinnerPl) & "!"
End Select

'Now let's calculate the final scores and put this player in the
'high score list, if they deserve it.
StatScreen.CalculateScores
HiScores.AddPlayerToHighScoreList (WinnerPl)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then
  'We're trying to close with the close gadget.
  
  Dim Resp
  
  'See if we've saved this map yet.  We want to save hi scores with the map.
  If MyMap.MapName <> "" Then
    'We've already saved the map, so update it.
    QuickMapUpdate
    LastMapPath = MyMap.MapName
    LastMapStamp = MyMap.MapStamp
    Land.SetupMenusForGameOver
    If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
    GameMode = GM_BUILDING_TITLE
    MyNetworkRole = NW_NONE
  Else
    'If we're configured to do so, prompt the user to save.
    If (CFGPrompt = 1) And (MyNetworkRole = NW_NONE) Then
      'We haven't saved this map yet, ask the user if they want to.
      Resp = MsgBox("This map has not been saved yet.  You must save the map" & vbCrLf & _
             "to keep a hi score list.  Would you like to save this map?", _
             vbInformation Or vbYesNoCancel, "Map Not Saved")
      If Resp = vbNo Then
        LastMapPath = MyMap.MapName
        LastMapStamp = MyMap.MapStamp
        Land.SetupMenusForGameOver
        If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
        GameMode = GM_BUILDING_TITLE
        MyNetworkRole = NW_NONE
      ElseIf Resp = vbYes Then
        Cancel = 1
        SaveMap
      ElseIf Resp = vbCancel Then
        Cancel = 1
      End If
    Else
      'No prompt, so just leave.
      LastMapPath = MyMap.MapName
      LastMapStamp = MyMap.MapStamp
      Land.SetupMenusForGameOver
      If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
      GameMode = GM_BUILDING_TITLE
      MyNetworkRole = NW_NONE
    End If
  End If
  
End If

End Sub

Private Sub HiScoreButton_Click()

HiScores.Show vbModal

End Sub

Private Sub StatsButton_Click()

StatScreen.Show vbModal

End Sub

Private Sub WinningButton_Click()

Dim Resp

'See if we've saved this map yet.  We want to save hi scores with the map.
If MyMap.MapName <> "" Then
  'We've already saved the map, so update it.
  QuickMapUpdate
  LastMapPath = MyMap.MapName
  LastMapStamp = MyMap.MapStamp
  Land.SetupMenusForGameOver
  If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
  GameMode = GM_BUILDING_TITLE
  MyNetworkRole = NW_NONE
  Unload Me
Else
  'If we're configured to do so, prompt the user to save.
  If (CFGPrompt = 1) And (MyNetworkRole = NW_NONE) Then
    'We haven't saved this map yet, ask the user if they want to.
    Resp = MsgBox("This map has not been saved yet.  You must save the map" & vbCrLf & _
           "to keep a hi score list.  Would you like to save this map?", _
           vbInformation Or vbYesNoCancel, "Map Not Saved")
    If Resp = vbNo Then
      LastMapPath = MyMap.MapName
      LastMapStamp = MyMap.MapStamp
      Land.SetupMenusForGameOver
      If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
      GameMode = GM_BUILDING_TITLE
      MyNetworkRole = NW_NONE
      Unload Me
    ElseIf Resp = vbYes Then
      SaveMap
    End If
  Else
    'No prompt, so just leave.
    LastMapPath = MyMap.MapName
    LastMapStamp = MyMap.MapStamp
    Land.SetupMenusForGameOver
    If MyNetworkRole <> NW_NONE Then NetworkForm.CancelBut_Click
    GameMode = GM_BUILDING_TITLE
    MyNetworkRole = NW_NONE
    Unload Me
  End If
End If

End Sub
