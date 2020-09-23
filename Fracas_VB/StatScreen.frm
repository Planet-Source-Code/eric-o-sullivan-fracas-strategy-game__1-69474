VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form StatScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Statistics"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   9210
   Icon            =   "StatScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9210
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKbutt 
      Caption         =   "OK"
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   6000
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "Rankings"
      Height          =   2055
      Left            =   5280
      TabIndex        =   10
      Top             =   4320
      Width           =   3015
      Begin MSFlexGridLib.MSFlexGrid FGrankings 
         Height          =   1695
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   4
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "HQs Defeated"
      Height          =   2055
      Left            =   5280
      TabIndex        =   8
      Top             =   0
      Width           =   3015
      Begin MSFlexGridLib.MSFlexGrid FGdefeated 
         Height          =   1695
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^|^|^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Totals"
      Height          =   2055
      Left            =   5280
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
      Begin MSFlexGridLib.MSFlexGrid FGtotals 
         Height          =   1695
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   4
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Troops Killed"
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid FGkilled 
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^|^|^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Countries Overtaken"
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid FGovertaken 
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^|^|^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Countries Attacked"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid FGattacked 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   8
         Cols            =   8
         AllowBigSelection=   0   'False
         GridLinesFixed  =   0
         ScrollBars      =   0
         Appearance      =   0
         FormatString    =   "^|^|^|^|^|^|^"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "StatScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Form_Load()

Dim i As Integer
Dim j As Integer

'Calculate everything fresh.
CalculateScores

'Setup table data.
FGattacked.Row = 0
FGattacked.Col = 0
FGattacked.ColWidth(0) = 62 * Screen.TwipsPerPixelX
FGattacked.CellBackColor = StatScreen.BackColor
FGattacked.Text = vbNullString
FGovertaken.Row = 0
FGovertaken.Col = 0
FGovertaken.ColWidth(0) = 62 * Screen.TwipsPerPixelX
FGovertaken.CellBackColor = StatScreen.BackColor
FGovertaken.Text = vbNullString
FGkilled.Row = 0
FGkilled.Col = 0
FGkilled.ColWidth(0) = 62 * Screen.TwipsPerPixelX
FGkilled.CellBackColor = StatScreen.BackColor
FGkilled.Text = vbNullString
FGdefeated.Row = 0
FGdefeated.Col = 0
FGdefeated.ColWidth(0) = 64 * Screen.TwipsPerPixelX
FGdefeated.CellBackColor = StatScreen.BackColor
FGdefeated.Text = vbNullString
FGtotals.Row = 0
FGtotals.Col = 0
FGtotals.ColWidth(0) = 62 * Screen.TwipsPerPixelX
FGtotals.CellBackColor = StatScreen.BackColor
FGtotals.Text = vbNullString
FGrankings.Row = 0
FGrankings.Col = 0
FGrankings.ColWidth(0) = 62 * Screen.TwipsPerPixelX
FGrankings.CellBackColor = StatScreen.BackColor
FGrankings.Text = vbNullString
'Set up player names, column widths, etc.
For i = 1 To 6
  FGattacked.Row = i
  FGattacked.Col = 0
  FGattacked.ColWidth(i) = 43 * Screen.TwipsPerPixelX
  FGovertaken.Row = i
  FGovertaken.Col = 0
  FGovertaken.ColWidth(i) = 43 * Screen.TwipsPerPixelX
  FGkilled.Row = i
  FGkilled.Col = 0
  FGkilled.ColWidth(i) = 43 * Screen.TwipsPerPixelX
  FGtotals.Row = i
  FGtotals.Col = 0
  FGrankings.Row = i
  FGrankings.Col = 0
  FGdefeated.Row = i
  FGdefeated.Col = 0
  FGdefeated.ColWidth(i) = 20.3 * Screen.TwipsPerPixelX
  If i < 3 Then
    FGtotals.ColWidth(i) = 43 * Screen.TwipsPerPixelX
    FGrankings.ColWidth(i) = 43 * Screen.TwipsPerPixelX
  End If
  
  If PlayerType(i) > PTYPE_INACTIVE Then
  
    FGattacked.Text = PlayerName(i)
    FGattacked.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 6
        FGattacked.Col = j
        FGattacked.CellForeColor = &H808080
      Next j
    End If
    FGattacked.Row = 0
    FGattacked.Col = i
    FGattacked.CellBackColor = PlayerColorCodes(Player(i))
    
    FGovertaken.Text = PlayerName(i)
    FGovertaken.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 6
        FGovertaken.Col = j
        FGovertaken.CellForeColor = &H808080
      Next j
    End If
    FGovertaken.Row = 0
    FGovertaken.Col = i
    FGovertaken.CellBackColor = PlayerColorCodes(Player(i))
    
    FGkilled.Text = PlayerName(i)
    FGkilled.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 6
        FGkilled.Col = j
        FGkilled.CellForeColor = &H808080
      Next j
    End If
    FGkilled.Row = 0
    FGkilled.Col = i
    FGkilled.CellBackColor = PlayerColorCodes(Player(i))
    
    FGtotals.Text = PlayerName(i)
    FGtotals.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 2
        FGtotals.Col = j
        FGtotals.CellForeColor = &H808080
      Next j
    End If
    
    FGrankings.Text = PlayerName(i)
    FGrankings.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 2
        FGrankings.Col = j
        FGrankings.CellForeColor = &H808080
      Next j
    End If
    
    FGdefeated.Text = PlayerName(i)
    FGdefeated.CellBackColor = PlayerColorCodes(Player(i))
    If NumOccupied(i) < 0 Then
      For j = 0 To 6
        FGdefeated.Col = j
        FGdefeated.CellForeColor = &H808080
      Next j
    End If
    FGdefeated.Row = 0
    FGdefeated.Col = i
    FGdefeated.CellBackColor = PlayerColorCodes(Player(i))
    
  Else
    For j = 0 To 6
      FGattacked.Row = j
      FGattacked.Col = i
      FGattacked.Text = vbNullString
      FGattacked.CellBackColor = StatScreen.BackColor
      FGattacked.Row = i
      FGattacked.Col = j
      FGattacked.Text = vbNullString
      FGattacked.CellBackColor = StatScreen.BackColor
      
      FGovertaken.Row = j
      FGovertaken.Col = i
      FGovertaken.Text = vbNullString
      FGovertaken.CellBackColor = StatScreen.BackColor
      FGovertaken.Row = i
      FGovertaken.Col = j
      FGovertaken.Text = vbNullString
      FGovertaken.CellBackColor = StatScreen.BackColor
      
      FGkilled.Row = j
      FGkilled.Col = i
      FGkilled.Text = vbNullString
      FGkilled.CellBackColor = StatScreen.BackColor
      FGkilled.Row = i
      FGkilled.Col = j
      FGkilled.Text = vbNullString
      FGkilled.CellBackColor = StatScreen.BackColor
      
      FGdefeated.Row = j
      FGdefeated.Col = i
      FGdefeated.Text = vbNullString
      FGdefeated.CellBackColor = StatScreen.BackColor
      FGdefeated.Row = i
      FGdefeated.Col = j
      FGdefeated.Text = vbNullString
      FGdefeated.CellBackColor = StatScreen.BackColor
    Next j
    
    FGtotals.Text = vbNullString
    FGtotals.CellBackColor = StatScreen.BackColor
    FGtotals.Col = 1
    FGtotals.Text = vbNullString
    FGtotals.CellBackColor = StatScreen.BackColor
    FGtotals.Col = 2
    FGtotals.Text = vbNullString
    FGtotals.CellBackColor = StatScreen.BackColor
    If i < 3 Then
      FGtotals.Row = 0
      FGtotals.Col = i
      FGtotals.CellBackColor = StatScreen.BackColor
    End If
    
    FGrankings.Text = vbNullString
    FGrankings.CellBackColor = StatScreen.BackColor
    FGrankings.Col = 1
    FGrankings.Text = vbNullString
    FGrankings.CellBackColor = StatScreen.BackColor
    FGrankings.Col = 2
    FGrankings.Text = vbNullString
    FGrankings.CellBackColor = StatScreen.BackColor
    If i < 3 Then
      FGrankings.Row = 0
      FGrankings.Col = i
      FGrankings.CellBackColor = StatScreen.BackColor
    End If
    
  End If
  'Diagonals will have no data.
  FGattacked.Row = i
  FGattacked.Col = i
  FGattacked.CellBackColor = StatScreen.BackColor
  FGovertaken.Row = i
  FGovertaken.Col = i
  FGovertaken.CellBackColor = StatScreen.BackColor
  FGkilled.Row = i
  FGkilled.Col = i
  FGkilled.CellBackColor = StatScreen.BackColor
  FGdefeated.Row = i
  FGdefeated.Col = i
  FGdefeated.CellBackColor = StatScreen.BackColor
Next i
'Set up headings on some of the smaller tables.
FGtotals.Row = 0
FGtotals.Col = 1
FGtotals.Text = "Countries"
FGtotals.Col = 2
FGtotals.Text = "Troops"
FGtotals.ColWidth(1) = 61 * Screen.TwipsPerPixelX
FGtotals.ColWidth(2) = 61 * Screen.TwipsPerPixelX
FGrankings.Row = 0
FGrankings.Col = 1
FGrankings.Text = "Score"
FGrankings.Col = 2
FGrankings.Text = "Rank"
FGrankings.ColWidth(1) = 61 * Screen.TwipsPerPixelX
FGrankings.ColWidth(2) = 61 * Screen.TwipsPerPixelX

'Assign the correct values from our statistics arrays.
For i = 1 To 6
  For j = 1 To 6
    FGattacked.Row = i
    FGattacked.Col = j
    FGovertaken.Row = i
    FGovertaken.Col = j
    FGkilled.Row = i
    FGkilled.Col = j
    'Countries attacked.
    If STATattacked(i, j) = 0 Then
      FGattacked.Text = vbNullString
    Else
      FGattacked.Text = Trim(Str(STATattacked(i, j)))
    End If
    'Countries overtaken.
    If STATovertaken(i, j) = 0 Then
      FGovertaken.Text = vbNullString
    Else
      FGovertaken.Text = Trim(Str(STATovertaken(i, j)))
    End If
    'Troops killed.
    If STATkilled(i, j) = 0 Then
      FGkilled.Text = vbNullString
    Else
      FGkilled.Text = Trim(Str(STATkilled(i, j)))
    End If
    'HQs defeated.
    FGdefeated.Row = i
    FGdefeated.Col = j
    If (i <> j) And (PlayerType(i) > PTYPE_INACTIVE) Then
      Select Case STATdefeated(i, j)
        Case -1:
          FGdefeated.Text = "R"
        Case 0:
          FGdefeated.Text = vbNullString
        Case 1:
          FGdefeated.Text = "X"
      End Select
    End If
  Next j
  'Totals.
  FGtotals.Row = i
  FGtotals.Col = 1
  If STATcountries(i) <= 0 Then
    FGtotals.Text = vbNullString
  Else
    FGtotals.Text = Trim(Str(STATcountries(i)))
  End If
  FGtotals.Col = 2
  If STATtroops(i) <= 0 Then
    FGtotals.Text = vbNullString
  Else
    FGtotals.Text = Trim(Str(STATtroops(i)))
  End If
  'Scores.
  FGrankings.Row = i
  FGrankings.Col = 1
  If STATscore(i) <= 0 Then
    'Not playing.
    FGrankings.Text = vbNullString
    FGrankings.Col = 2
    If PlayerType(i) = PTYPE_INACTIVE Then
      FGrankings.Text = vbNullString
    Else
      FGrankings.Text = Trim(Str(STATrank(i)))
    End If
  Else
    FGrankings.Text = Trim(Str(STATscore(i)))
    If STATrank(i) = 1 Then
      'This player is in first place!  Highlight.
      FGrankings.CellFontBold = True
    Else
      FGrankings.CellFontBold = False
    End If
    FGrankings.Col = 2
    FGrankings.Text = Trim(Str(STATrank(i)))
    If STATrank(i) = 1 Then
      'This player is in first place!  Highlight.
      FGrankings.CellFontBold = True
    Else
      FGrankings.CellFontBold = False
    End If
  End If
  
Next i

'Set cursors to non-visible cells.
FGattacked.Row = 7
FGattacked.Col = 7
FGovertaken.Row = 7
FGovertaken.Col = 7
FGkilled.Row = 7
FGkilled.Col = 7
FGdefeated.Row = 7
FGdefeated.Col = 7
FGtotals.Row = 7
FGtotals.Col = 3
FGrankings.Row = 7
FGrankings.Col = 3

End Sub

Public Sub ResetAllStats()

Dim i As Integer
Dim j As Integer

'This sub completely zeroes out all statistical information.
For i = 1 To MAX_PLAYERS
  For j = 1 To MAX_PLAYERS
    STATattacked(i, j) = 0
    STATovertaken(i, j) = 0
    STATkilled(i, j) = 0
    STATdefeated(i, j) = 0
  Next j
  STATscore(i) = 0
Next i

End Sub

Public Sub UpdateAttackedStats(Attacker As Integer, Attackee As Integer)

'This sub increments the Country Attacked stat for the passed players.
STATattacked(Attacker, Attackee) = STATattacked(Attacker, Attackee) + 1

End Sub

Public Sub UpdateOvertakenStats(Attacker As Integer, Attackee As Integer)

'This sub increments the Country Overtaken stat for the passed players.
STATovertaken(Attacker, Attackee) = STATovertaken(Attacker, Attackee) + 1

End Sub

Public Sub UpdateKilledStats(Attacker As Integer, Attackee As Integer, Amount As Long)

'This sub increments the Troops Killed stat for the passed players.
STATkilled(Attacker, Attackee) = STATkilled(Attacker, Attackee) + Amount

End Sub

Public Sub UpdateDefeatedStats(Attacker As Integer, Attackee As Integer)

'This sub increments the HQ Overtaken stat for the passed players.
If (Attacker <> Attackee) Then STATdefeated(Attacker, Attackee) = 1

End Sub

Public Sub UpdateResignedStats(Resignee As Integer)

'This sub increments the HQ Overtaken stat for the passed players.

Dim i As Integer

For i = 1 To MAX_PLAYERS
  If i <> Resignee Then STATdefeated(i, Resignee) = -1
Next i

End Sub

Private Sub OKbutt_Click()

Unload Me

End Sub

Public Sub CalculateScores()

Dim i As Integer
Dim j As Integer
Dim k As Long
Dim AttackedSubTotal(MAX_PLAYERS) As Long
Dim OvertakenSubTotal(MAX_PLAYERS) As Long
Dim KilledSubTotal(MAX_PLAYERS) As Long
Dim DefeatedSubTotal(MAX_PLAYERS) As Long
Dim ResignedSubTotal(MAX_PLAYERS) As Long

Dim TempPlayerNum(MAX_PLAYERS) As Integer
Dim TempScore(MAX_PLAYERS) As Long
Dim RankChanged As Boolean

Dim AttackedTotal As Long
Dim OvertakenTotal As Long
Dim KilledTotal As Long
Dim CountryTotal As Long
Dim TroopTotal As Long

Dim PointsPerHQ As Single

NumPlayers = Players.CalcNumPlayers
AttackedTotal = 0
OvertakenTotal = 0
KilledTotal = 0
CountryTotal = 0
TroopTotal = 0
'First, calculate everyone's score.
For i = 1 To MAX_PLAYERS
  'Zero this player's subtotals out.
  AttackedSubTotal(i) = 0
  OvertakenSubTotal(i) = 0
  KilledSubTotal(i) = 0
  DefeatedSubTotal(i) = 0
  ResignedSubTotal(i) = 0
  STATcountries(i) = 0
  STATtroops(i) = 0
  'Calculate this player's score for warfare.
  For j = 1 To MAX_PLAYERS
    AttackedSubTotal(i) = AttackedSubTotal(i) + STATattacked(i, j)
    OvertakenSubTotal(i) = OvertakenSubTotal(i) + STATovertaken(i, j)
    KilledSubTotal(i) = KilledSubTotal(i) + STATkilled(i, j)
    If STATdefeated(i, j) = 1 Then
      DefeatedSubTotal(i) = DefeatedSubTotal(i) + 1
    ElseIf STATdefeated(i, j) = -1 Then
      ResignedSubTotal(i) = ResignedSubTotal(i) + 1
    End If
  Next j
  'We get the troop and country subtotals by looking at the MAP.
  For k = 1 To MyMap.NumberOfCountries
    If MyMap.Owner(k) = i Then
      'This player owns this country.  Count it for scoring.
      'Note that unowned countries don't contribute in any way to score.
      STATcountries(i) = STATcountries(i) + 1
      CountryTotal = CountryTotal + 1
      STATtroops(i) = STATtroops(i) + MyMap.TroopCount(k)
      TroopTotal = TroopTotal + MyMap.TroopCount(k)
    End If
  Next k
Next i

'Now get absolute totals for each value.
For i = 1 To MAX_PLAYERS
  AttackedTotal = AttackedTotal + AttackedSubTotal(i)
  OvertakenTotal = OvertakenTotal + OvertakenSubTotal(i)
  KilledTotal = KilledTotal + KilledSubTotal(i)
Next i
'Now divide to get a score for each between 0 and 100.  There are 100
'points to distribute in each category!
For i = 1 To MAX_PLAYERS
  If PlayerType(i) = PTYPE_INACTIVE Then
    'Not playing - no score.
    STATscore(i) = -1
  Else
    STATscore(i) = 0
    If AttackedTotal > 0 Then
      STATscore(i) = STATscore(i) + (100 * AttackedSubTotal(i) / AttackedTotal)
    End If
    If OvertakenTotal > 0 Then
      STATscore(i) = STATscore(i) + (100 * OvertakenSubTotal(i) / OvertakenTotal)
    End If
    If KilledTotal > 0 Then
      STATscore(i) = STATscore(i) + (100 * KilledSubTotal(i) / KilledTotal)
    End If
    'Jason -- redo these, look at map variables.
    If CountryTotal > 0 Then
      STATscore(i) = STATscore(i) + (100 * STATcountries(i) / CountryTotal)
    End If
    If TroopTotal > 0 Then
      STATscore(i) = STATscore(i) + (100 * STATtroops(i) / TroopTotal)
    End If
    'Determine how much each HQ is worth.  There are 500 points unaccounted for
    'out of the 1000, so we need to do some division.
    'Thus, with 2 players, one of them will get the whole 500.
    'With 6 players, each of the other five is worth 100.
    PointsPerHQ = 500 / (NumPlayers - 1)
    'Add in points per HQ they've killed...
    STATscore(i) = Int(STATscore(i) + (PointsPerHQ * DefeatedSubTotal(i)))
    'Plus points for someone resigning...
    STATscore(i) = Int(STATscore(i) + ((PointsPerHQ / (NumPlayers - 1)) * ResignedSubTotal(i)))
  End If
Next i

'Now order the scores and assign ranks.
For i = 1 To MAX_PLAYERS
  TempPlayerNum(i) = i
  TempScore(i) = STATscore(i)
Next i

'Bubble sort the six.
Do
  RankChanged = False
  For i = 1 To (MAX_PLAYERS - 1)
    If TempScore(i) < TempScore(i + 1) Then
      'This score is less than the one below it, so swap the two.
      k = TempScore(i)
      TempScore(i) = TempScore(i + 1)
      TempScore(i + 1) = k
      'Swap the score *and* this player's number.
      j = TempPlayerNum(i)
      TempPlayerNum(i) = TempPlayerNum(i + 1)
      TempPlayerNum(i + 1) = j
      'Mark as changed so we do another iteration.
      RankChanged = True
    End If
  Next i
Loop Until RankChanged = False

'Now determine rank according to the player number order.
For i = 1 To MAX_PLAYERS
  STATrank(TempPlayerNum(i)) = i
Next i

'If there is a tie, then two players are tied at the same rank.
For i = 2 To MAX_PLAYERS
  If TempScore(i) = TempScore(i - 1) Then
    'This player is tied with the person before them.
    STATrank(TempPlayerNum(i)) = STATrank(TempPlayerNum(i - 1))
  End If
Next i

End Sub
