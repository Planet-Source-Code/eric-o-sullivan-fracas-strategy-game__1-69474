VERSION 5.00
Begin VB.Form Players 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Player Selection"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Players.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   6
      Left            =   3630
      Style           =   2  'Dropdown List
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1300
   End
   Begin VB.TextBox Pname 
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   6
      Text            =   "Player 6"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   5
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1300
   End
   Begin VB.TextBox Pname 
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Text            =   "Player 5"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Player6 
      Caption         =   "Player 6"
      Height          =   2895
      Left            =   3480
      TabIndex        =   37
      Top             =   3120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   49
         Top             =   1065
         Width           =   1185
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   6
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2040
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Player5 
      Caption         =   "Player 5"
      Height          =   2895
      Left            =   1800
      TabIndex        =   32
      Top             =   3120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   48
         Top             =   1065
         Width           =   1215
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2040
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   5
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.TextBox Pname 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Text            =   "Player 4"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Pname 
      Height          =   285
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Text            =   "Player 3"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Pname 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Text            =   "Player 2"
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton PlayerCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton PlayerOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   6240
      Width           =   1095
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   4
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1300
   End
   Begin VB.Frame Player4 
      Caption         =   "Player 4"
      Height          =   2895
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   1065
         Width           =   1260
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   4
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2040
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   3
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1300
   End
   Begin VB.Frame Player3 
      Caption         =   "Player 3"
      Height          =   2895
      Left            =   3480
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   1065
         Width           =   1230
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   3
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2040
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   2
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1300
   End
   Begin VB.Frame Player2 
      Caption         =   "Player 2"
      Height          =   2895
      Left            =   1800
      TabIndex        =   13
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   45
         Top             =   1065
         Width           =   1200
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   2
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Index           =   1
      ItemData        =   "Players.frx":014A
      Left            =   270
      List            =   "Players.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1300
   End
   Begin VB.Frame Player1 
      Caption         =   "Player 1"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Network 
         Caption         =   "Network"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1065
         Width           =   1215
      End
      Begin VB.ComboBox Personal 
         Height          =   315
         Index           =   1
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Pname 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Text            =   "Player 1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton None 
         Caption         =   "Inactive"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton Computer 
         Caption         =   "Computer"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Human 
         Caption         =   "Human"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   690
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Players"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Color_Click(Index As Integer)

UpdateColorDropdowns
If Players.Visible = True Then
  PlayerOK.SetFocus
End If

End Sub

Private Sub UpdateColorDropdowns()

'This sub just writes the chosen colors to the BackColor of the
'color drop-downs.

Dim i As Integer
Dim j As Integer

For i = 1 To 6
  For j = 1 To 11
    If Color(i).Text = PlayerColors(j) Then
      Color(i).BackColor = (PlayerColorCodes(j))
      Color(i).ForeColor = PlayerTextColor(j)
    End If
  Next j
Next i

End Sub

Private Sub Computer_Click(Index As Integer)

Personal(Index).Enabled = True

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

For i = 1 To 6
  'Set up the color control.
  Color(i).Clear
  For j = 1 To 11  'Not 12 since white is reserved for flashing.
    'Don't add the unoccupied color to the dropdowns.
    If (j <> CFGUnoccupiedColor) Then Color(i).AddItem PlayerColors(j)
  Next j
  Color(i).Text = PlayerColors(Player(i))
  'Set up the player types.
  Select Case (PlayerType(i))
    Case PTYPE_INACTIVE:  None(i).Value = True
    Case PTYPE_HUMAN:  Human(i).Value = True
    Case PTYPE_COMPUTER:  Computer(i).Value = True
    Case PTYPE_NETWORK:  Network(i).Value = True
  End Select
  'Set up the player names.
  Pname(i).Text = PlayerName(i)
  'Set up the computer personalities.
  For j = 1 To NUMPERSONALITIES
    Personal(i).AddItem (PersonalityName(j))
  Next j
  Personal(i).Text = PersonalityName(Personality(i))
  Personal(i).Enabled = False
  If PlayerType(i) = PTYPE_COMPUTER Then Personal(i).Enabled = True
Next i

UpdateColorDropdowns

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then
  PlayerCancel_Click
End If

End Sub

Private Sub Human_Click(Index As Integer)

Personal(Index).Enabled = False

End Sub

Private Sub Network_Click(Index As Integer)

Personal(Index).Enabled = False

End Sub

Private Sub None_Click(Index As Integer)

Personal(Index).Enabled = False

End Sub

Private Sub PlayerCancel_Click()

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Make this settings form disappear.
Unload Me

End Sub

Private Sub PlayerOK_Click()

Dim Numpl As Integer
Dim i As Integer
Dim j As Integer

Numpl = 0

'Check to make sure we have at least two players.
For i = 1 To 6
  If None(i).Value = False Then
    Numpl = Numpl + 1
  End If
Next i
If Numpl < 2 Then
  MsgBox ("There must be at least two players.")
  Exit Sub
End If

'Check to make sure we have all different colors.  Inactive players
'won't count during this check.
For i = 1 To 6
  For j = 1 To 6
    If (Color(i).Text = Color(j).Text) And (i <> j) And (None(i).Value = False) And (None(j).Value = False) Then
      MsgBox ("Each active player must have a unique color.")
      Exit Sub
    End If
  Next j
Next i

'Check to make sure we have valid player names.
For i = 1 To 6
  'Names can't be null.
  If (Pname(i).Text = "") And (None(i).Value = False) Then
    MsgBox ("Each active player must have a name.")
    Exit Sub
  End If
  'Names can't be long.
  If (Len(Pname(i).Text) > 10) And (None(i).Value = False) Then
    MsgBox ("Player names are limited to 10 characters.")
    Exit Sub
  End If
  'Names can't have commas or quotes because it screws up the saved games.
  If (InStr(Pname(i).Text, ",") > 0) Or (InStr(Pname(i).Text, Chr$(34)) > 0) Then
    MsgBox ("One or more player names contain invalid characters.")
    Exit Sub
  End If
  'Names can't be the same.
  For j = 1 To 6
    If (i <> j) And (Pname(i) = Pname(j)) Then
      MsgBox ("Player names must be unique.")
      Exit Sub
    End If
  Next j
Next i

'Copy the form settings into our permanent variables.
For i = 1 To 6
  'Put the proper color into the Player array.
  For j = 1 To 11
    If Color(i).Text = PlayerColors(j) Then
      Player(i) = j
    End If
  Next j
  'Put the correct type into the type array.
  If None(i).Value = True Then PlayerType(i) = PTYPE_INACTIVE
  If Human(i).Value = True Then PlayerType(i) = PTYPE_HUMAN
  If Computer(i).Value = True Then PlayerType(i) = PTYPE_COMPUTER
  If Network(i).Value = True Then PlayerType(i) = PTYPE_NETWORK
  'Record the player name.
  PlayerName(i) = Pname(i).Text
  Land.Menu1st(i).Caption = Pname(i).Text
  'Put the proper personality into the Personality array.
  For j = 1 To NUMPERSONALITIES
    If Personal(i).Text = PersonalityName(j) Then
      Personality(i) = j
    End If
  Next j
Next i

'Get the number of players.
NumPlayers = Numpl

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Make this settings form disappear.
Unload Me

End Sub

Public Function CalcNumPlayers() As Integer

Dim i As Integer

CalcNumPlayers = 0
For i = 1 To MAX_PLAYERS
  If PlayerType(i) > PTYPE_INACTIVE Then CalcNumPlayers = CalcNumPlayers + 1
Next i

End Function
