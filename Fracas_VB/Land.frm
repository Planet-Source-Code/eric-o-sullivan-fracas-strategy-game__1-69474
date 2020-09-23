VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Land 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fracas"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10890
   FillStyle       =   0  'Solid
   Icon            =   "Land.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer FlashTimer 
      Interval        =   1
      Left            =   10320
      Top             =   6840
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4320
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton ResignBut 
      Caption         =   "Resign..."
      Height          =   375
      Left            =   9600
      TabIndex        =   37
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox TroopMoveAll 
      Height          =   255
      Left            =   9120
      Picture         =   "Land.frx":0442
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   36
      Top             =   4200
      Width           =   255
   End
   Begin VB.PictureBox TroopMove3Qt 
      Height          =   255
      Left            =   8760
      Picture         =   "Land.frx":0954
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin VB.PictureBox TroopMoveHlf 
      Height          =   255
      Left            =   8400
      Picture         =   "Land.frx":0E66
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   34
      Top             =   4200
      Width           =   255
   End
   Begin VB.PictureBox TroopMoveQtr 
      Height          =   255
      Left            =   8040
      Picture         =   "Land.frx":1378
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   33
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton CancelBut 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9600
      TabIndex        =   31
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton TroopMoveDn 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9150
      TabIndex        =   30
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton TroopMoveUp 
      Caption         =   "+"
      Height          =   255
      Left            =   9150
      TabIndex        =   29
      Top             =   3600
      Width           =   255
   End
   Begin VB.TextBox TroopMoveNum 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8310
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3720
      Width           =   735
   End
   Begin VB.Frame CntryFrame 
      Height          =   2655
      Left            =   9600
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label DefendWtr 
         Alignment       =   2  'Center
         Caption         =   "Water: 0"
         Height          =   255
         Left            =   75
         TabIndex        =   27
         Top             =   2375
         Width           =   945
      End
      Begin VB.Label DefendLnd 
         Alignment       =   2  'Center
         Caption         =   "Land: 0"
         Height          =   255
         Left            =   75
         TabIndex        =   26
         Top             =   2150
         Width           =   945
      End
      Begin VB.Label DefendTot 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   25
         Top             =   1895
         Width           =   945
      End
      Begin VB.Label DefendTxt 
         Alignment       =   2  'Center
         Caption         =   "Defend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label AttackWtr 
         Alignment       =   2  'Center
         Caption         =   "Water: 0"
         Height          =   255
         Left            =   75
         TabIndex        =   23
         Top             =   1380
         Width           =   945
      End
      Begin VB.Label AttackLnd 
         Alignment       =   2  'Center
         Caption         =   "Land: 0"
         Height          =   255
         Left            =   75
         TabIndex        =   22
         Top             =   1155
         Width           =   945
      End
      Begin VB.Label AttackTot 
         Alignment       =   2  'Center
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   21
         Top             =   900
         Width           =   945
      End
      Begin VB.Label AttackTxt 
         Alignment       =   2  'Center
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   705
         Width           =   855
      End
      Begin VB.Label CntryName 
         Alignment       =   2  'Center
         Caption         =   "CntryName"
         Height          =   435
         Left            =   90
         TabIndex        =   19
         Top             =   195
         Width           =   945
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton PassTurnBut 
      Caption         =   "Pass..."
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer GameTimer 
      Interval        =   1
      Left            =   9840
      Top             =   6840
   End
   Begin VB.PictureBox WaterPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   1320
      Picture         =   "Land.frx":188A
      ScaleHeight     =   300
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox MapBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7335
      Left            =   6720
      ScaleHeight     =   7275
      ScaleWidth      =   9795
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.PictureBox BallPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1635
      Left            =   120
      Picture         =   "Land.frx":1EFC
      ScaleHeight     =   1575
      ScaleWidth      =   1050
      TabIndex        =   4
      Top             =   7320
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.Timer BallTimer 
      Interval        =   1
      Left            =   9360
      Top             =   6840
   End
   Begin VB.PictureBox PicBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7335
      Left            =   3600
      ScaleHeight     =   7275
      ScaleWidth      =   9795
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.PictureBox Ocean 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7140
      Left            =   1440
      ScaleHeight     =   7080
      ScaleWidth      =   9720
      TabIndex        =   2
      Top             =   8160
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.PictureBox LandMap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   120
      Picture         =   "Land.frx":40C6
      ScaleHeight     =   1920
      ScaleWidth      =   9120
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   9180
   End
   Begin VB.Frame PlyrFrame 
      Height          =   3045
      Left            =   8400
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals6"
         Height          =   255
         Index           =   6
         Left            =   75
         TabIndex        =   41
         Top             =   2755
         Width           =   945
      End
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals5"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   40
         Top             =   2285
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P6Name"
         Height          =   255
         Index           =   6
         Left            =   75
         TabIndex        =   39
         Top             =   2530
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P5Name"
         Height          =   255
         Index           =   5
         Left            =   75
         TabIndex        =   38
         Top             =   2060
         Width           =   945
      End
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals4"
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   17
         Top             =   1815
         Width           =   945
      End
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals3"
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   16
         Top             =   1345
         Width           =   945
      End
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals2"
         Height          =   255
         Index           =   2
         Left            =   75
         TabIndex        =   15
         Top             =   875
         Width           =   945
      End
      Begin VB.Label PlyrTotals 
         Alignment       =   2  'Center
         Caption         =   "PlyrTotals1"
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   405
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P1Name"
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   180
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P2Name"
         Height          =   255
         Index           =   2
         Left            =   75
         TabIndex        =   12
         Top             =   650
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P3Name"
         Height          =   255
         Index           =   3
         Left            =   75
         TabIndex        =   11
         Top             =   1120
         Width           =   945
      End
      Begin VB.Label PlyrNames 
         Alignment       =   2  'Center
         Caption         =   "P4Name"
         Height          =   255
         Index           =   4
         Left            =   75
         TabIndex        =   10
         Top             =   1590
         Width           =   945
      End
   End
   Begin VB.Label TroopMoveLbl 
      Caption         =   "Troops :"
      Height          =   255
      Left            =   8310
      TabIndex        =   32
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Oopsie 
      Caption         =   "Oopsie Text."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   9135
   End
   Begin VB.Label Commentary 
      BackColor       =   &H80000004&
      Caption         =   "Running Commentary Text."
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   9135
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuSave 
         Caption         =   "Save Map"
      End
      Begin VB.Menu MenuSaveGame 
         Caption         =   "Save Game"
      End
      Begin VB.Menu MenuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuGame 
      Caption         =   "Game"
      Begin VB.Menu MenuNew 
         Caption         =   "New Game, New Map"
      End
      Begin VB.Menu MenuSame 
         Caption         =   "New Game, Same Map"
      End
      Begin VB.Menu MenuLoad 
         Caption         =   "New Game, Load Map"
      End
      Begin VB.Menu MenuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu MenuLoadSavedGame 
         Caption         =   "Load Saved Game"
      End
      Begin VB.Menu MenuBar18 
         Caption         =   "-"
      End
      Begin VB.Menu MenuJoinGame 
         Caption         =   "Join A Network Game"
      End
      Begin VB.Menu MenuBar188 
         Caption         =   "-"
      End
      Begin VB.Menu MenuChat 
         Caption         =   "Chat..."
      End
      Begin VB.Menu Menubar66 
         Caption         =   "-"
      End
      Begin VB.Menu MenuGetOptionsFromMapFile 
         Caption         =   "Get Options From Map File"
      End
      Begin VB.Menu Menubar6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStats 
         Caption         =   "Statistics..."
      End
      Begin VB.Menu MenuHiScores 
         Caption         =   "Hi Scores..."
      End
      Begin VB.Menu MenuBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu MenuAbortGame 
         Caption         =   "Abort Game"
      End
   End
   Begin VB.Menu MenuOpts 
      Caption         =   "Options"
      Begin VB.Menu MenuPlayers 
         Caption         =   "Players..."
      End
      Begin VB.Menu Menubar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuTroopPl 
         Caption         =   "Initial Troop Placement"
         Begin VB.Menu MenuInitTroops 
            Caption         =   "None"
            Index           =   1
         End
         Begin VB.Menu MenuInitTroops 
            Caption         =   "Sparse"
            Index           =   2
         End
         Begin VB.Menu MenuInitTroops 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuInitTroops 
            Caption         =   "Dense"
            Index           =   4
         End
         Begin VB.Menu MenuInitTroops 
            Caption         =   "All Terrain Inhabited"
            Index           =   5
         End
      End
      Begin VB.Menu MenuInitCts 
         Caption         =   "Initial Troop Counts"
         Begin VB.Menu MenuInitTroopCts 
            Caption         =   "Few"
            Index           =   1
         End
         Begin VB.Menu MenuInitTroopCts 
            Caption         =   "Normal"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuInitTroopCts 
            Caption         =   "Lots"
            Index           =   3
         End
      End
      Begin VB.Menu MenuBonusTroops 
         Caption         =   "Bonus Troops (Per Country Per Turn)"
         Begin VB.Menu MenuBonus 
            Caption         =   "Skip Reinforcement Phase"
            Index           =   0
         End
         Begin VB.Menu MenuBonus 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu MenuBonus 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu MenuBonus 
            Caption         =   "3  -- Normal"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuBonus 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu MenuBonus 
            Caption         =   "5 -- Lots"
            Index           =   5
         End
      End
      Begin VB.Menu menuFriendly 
         Caption         =   "Naval Support"
         Begin VB.Menu MenuShips 
            Caption         =   "Rafts (1/10 of country strength available over water)"
            Index           =   1
         End
         Begin VB.Menu MenuShips 
            Caption         =   "Skiffs (1/4 of country strength available over water)"
            Index           =   2
         End
         Begin VB.Menu MenuShips 
            Caption         =   "Destroyers (1/2 of country strength over water-- Normal)"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuShips 
            Caption         =   "Battleships (Full maritime offensive/defensive capabilities)"
            Index           =   4
         End
      End
      Begin VB.Menu menuEnemy 
         Caption         =   "Enemy Ports"
         Begin VB.Menu MenuPorts 
            Caption         =   "Captured"
            Index           =   1
         End
         Begin VB.Menu MenuPorts 
            Caption         =   "Destroyed"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuPorts 
            Caption         =   "Depends On Casualties"
            Index           =   3
         End
         Begin VB.Menu MenuPorts 
            Caption         =   "Random"
            Index           =   4
         End
      End
      Begin VB.Menu menuEnemyDead 
         Caption         =   "When an enemy is conquered..."
         Begin VB.Menu MenuConquer 
            Caption         =   "...the victor claims all remaining countries (Napoleonic Wars)"
            Index           =   1
         End
         Begin VB.Menu MenuConquer 
            Caption         =   "...the remaining countries become neutral (WW I)"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuConquer 
            Caption         =   "...the remaining countries stay loyal (WW II)"
            Index           =   3
         End
         Begin VB.Menu MenuConquer 
            Caption         =   "...the enemy is completely eradicated (WW III)"
            Index           =   4
         End
         Begin VB.Menu MenuConquer 
            Caption         =   "...chaos erupts in the defeated countries"
            Index           =   5
         End
      End
      Begin VB.Menu MenuEvts 
         Caption         =   "Random Events"
         Begin VB.Menu MenuRandom 
            Caption         =   "Off"
            Index           =   1
         End
         Begin VB.Menu MenuRandom 
            Caption         =   "Infrequent"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuRandom 
            Caption         =   "Frequent"
            Index           =   3
         End
      End
      Begin VB.Menu MenuHQSelection 
         Caption         =   "HQ Selection"
         Begin VB.Menu MenuHQSelect 
            Caption         =   "Manual"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu MenuHQSelect 
            Caption         =   "Automatic"
            Index           =   2
         End
      End
      Begin VB.Menu MenuFirst 
         Caption         =   "First Turn"
         Begin VB.Menu Menu1st 
            Caption         =   "Player 1"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Player 2"
            Index           =   2
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Player 3"
            Index           =   3
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Player 4"
            Index           =   4
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Player 5"
            Index           =   5
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Player 6"
            Index           =   6
         End
         Begin VB.Menu Menu1st 
            Caption         =   "Random"
            Index           =   7
         End
      End
   End
   Begin VB.Menu MenuOptions 
      Caption         =   "Terraform"
      Begin VB.Menu MenuCountrySize 
         Caption         =   "Country Size"
         Begin VB.Menu MenuSize 
            Caption         =   "Tiny Independently-Owned Countries"
            Index           =   1
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Small Third-World Countries"
            Index           =   2
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Medium OPEC Countries"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Large First-World Countries"
            Index           =   4
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Continents"
            Index           =   5
         End
         Begin VB.Menu MenuSize 
            Caption         =   "Hodge-Podge"
            Index           =   6
         End
      End
      Begin VB.Menu MenuCountrySizeProp 
         Caption         =   "Country Size Proportions"
         Begin VB.Menu MenuProp 
            Caption         =   "Proportional Countries"
            Index           =   1
         End
         Begin VB.Menu MenuProp 
            Caption         =   "Somewhat Proportional Countries"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuProp 
            Caption         =   "Unproportional Countries"
            Index           =   3
         End
      End
      Begin VB.Menu MenuCountryShapes 
         Caption         =   "Country Shapes"
         Begin VB.Menu MenuShape 
            Caption         =   "Normal (no artificial distortion)"
            Index           =   1
         End
         Begin VB.Menu MenuShape 
            Caption         =   "Irregular"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuShape 
            Caption         =   "Very Irregular"
            Index           =   3
         End
      End
      Begin VB.Menu MenuMinLakeSize 
         Caption         =   "Minimum Allowed Lake Size"
         Begin VB.Menu MenuLakeSize 
            Caption         =   "No Lake Correction"
            Index           =   1
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Tiny Lakes"
            Index           =   2
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Medium Lakes"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuLakeSize 
            Caption         =   "Large Lakes"
            Index           =   4
         End
      End
      Begin VB.Menu MenuGlobal 
         Caption         =   "Approximate Land:Water ratio"
         Begin VB.Menu MenuPct 
            Caption         =   "1:9"
            Index           =   1
         End
         Begin VB.Menu MenuPct 
            Caption         =   "1:3"
            Index           =   2
         End
         Begin VB.Menu MenuPct 
            Caption         =   "1:1"
            Index           =   3
         End
         Begin VB.Menu MenuPct 
            Caption         =   "3:1"
            Checked         =   -1  'True
            Index           =   4
         End
         Begin VB.Menu MenuPct 
            Caption         =   "9:1"
            Index           =   5
         End
      End
      Begin VB.Menu MenuIsle 
         Caption         =   "Islands"
         Begin VB.Menu MenuIslands 
            Caption         =   "No Islands"
            Index           =   1
         End
         Begin VB.Menu MenuIslands 
            Caption         =   "Some Islands"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuIslands 
            Caption         =   "Lots of Islands"
            Index           =   3
         End
      End
      Begin VB.Menu MenuRes 
         Caption         =   "Window Size"
         Begin VB.Menu MenuResolution 
            Caption         =   "640x480"
            Index           =   1
         End
         Begin VB.Menu MenuResolution 
            Caption         =   "800x600"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu MenuResolution 
            Caption         =   "1024x768"
            Index           =   3
         End
      End
   End
   Begin VB.Menu MenuDraw 
      Caption         =   "Preferences"
      Begin VB.Menu MenuComputerSpeed 
         Caption         =   "Computer Player Speed"
         Begin VB.Menu MenuAISpeed 
            Caption         =   "Fastest"
            Index           =   1
         End
         Begin VB.Menu MenuAISpeed 
            Caption         =   "Fast"
            Index           =   2
         End
         Begin VB.Menu MenuAISpeed 
            Caption         =   "Medium"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu MenuAISpeed 
            Caption         =   "Slow"
            Index           =   4
         End
      End
      Begin VB.Menu MenuBorderChoice 
         Caption         =   "Borders"
         Begin VB.Menu MenuBorders 
            Caption         =   "Big Dotted Lines"
            Index           =   1
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Big Dots"
            Index           =   2
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Hash Marks"
            Index           =   3
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Small Dots"
            Index           =   4
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "Contours"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu MenuBorders 
            Caption         =   "None"
            Index           =   6
         End
      End
      Begin VB.Menu MenuSounds 
         Caption         =   "Sound Effects"
         Begin VB.Menu MenuSfx 
            Caption         =   "On"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu MenuSfx 
            Caption         =   "Off"
            Index           =   2
         End
      End
      Begin VB.Menu MenuAniSpeed 
         Caption         =   "Graphics Options..."
      End
      Begin VB.Menu MenuRefresh 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MenuTextFile 
         Caption         =   "View Help File"
      End
   End
End
Attribute VB_Name = "Land"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Fracas 2.0 BETA by Jason Merlo
'jmerlo@austin.rr.com
'jason.merlo@frco.com
'http://www.smozzie.com

Option Explicit
Option Base 1

Dim CadenceTime As Date
Dim CfgNumCountries As Integer

Dim Turn As Integer
Dim Phase As Integer

Public Sub SetTurn(Val As Integer)

Turn = Val

End Sub

Public Function GetTurn() As Integer

GetTurn = Turn

End Function

Public Sub SetPhase(Val As Integer)

Phase = Val

End Sub

Public Function GetPhase() As Integer

GetPhase = Phase

End Function

Private Sub Form_Load()

'If we already have an instance running, then LEAVE.
'If App.PrevInstance Then End   'TODO: Jason, remove this comment!

'SavePicture Me.LandMap.Picture, (App.Path + "\LandMap.bmp")
'End


'TODO: Jason remove this debug stuff!
'DebugForm.Show vbModeless

Randomize

MapMade = False
GameMode = GM_BUILDING_TITLE
MyNetworkRole = NW_NONE

SetupPaths
DetermineMaxRes

'No balls currently onscreen.
'Clean out the balls and waves arrays.
Call ClearAllBalls
Call ClearAllWaves

'Set up water colors array.
WaterColors(1) = &HEF1010
WaterColors(2) = &HFF2020
WaterColors(3) = &HFF4040
WaterColors(4) = &HFF0000
WaterColors(5) = &HE00000
WaterColors(6) = &HD00000

'Initialize a bunch of stuff.
Call InitPlayerData
Call InitNameStuff
Call InitPersonalities

'Collect all menu settings.
Call CollectMapSettings(False)
Call CollectMenuSettings(False)
Call CollectDrawSettings

'New in 2.0 BETA: an .ini file which contains all of our config parameters.
'Check it now, and if it's there, overwrite settings with what's in there.
ReadINI INIpath
FixResolution
SetupMenusForGameOver

'Create a new ocean pic.
Call CreateOcean

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'No matter who kills us or how we die, we need to try to gracefully
'kill our network connection.

If (MyNetworkRole = NW_CLIENT) Or (MyNetworkRole = NW_SERVER) Then
  NetworkForm.CancelBut_Click
End If

End Sub

Private Sub MenuAbortGame_Click()

Dim Resp As Long
Dim AbortStr As String

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

AbortStr = "Are you sure you want to abort the game?"
If MyNetworkRole = NW_SERVER Then
  AbortStr = AbortStr & vbCrLf & vbCrLf & _
    "WARNING!  You are the host of a networked game." & vbCrLf & _
    "Aborting will end the game for all networked players."
ElseIf MyNetworkRole = NW_CLIENT Then
  AbortStr = AbortStr & vbCrLf & vbCrLf & _
    "WARNING!  You are currently part of a networked game." & vbCrLf & _
    "You will not be able to enter this game again."
End If

Resp = MsgBox(AbortStr, vbYesNo, "Abort Game")
If Resp = vbNo Then Exit Sub

If MyNetworkRole = NW_CLIENT Or MyNetworkRole = NW_SERVER Then
  NetworkForm.CancelBut_Click
End If
LastMapPath = MyMap.MapName
LastMapStamp = MyMap.MapStamp
GameMode = GM_BUILDING_TITLE
MyNetworkRole = NW_NONE
SetupMenusForGameOver
Commentary.Caption = "Game aborted."
Oopsie.Caption = ""
Me.Refresh
Call DrawMap

End Sub

Private Function SetupNetwork() As Integer

'This function is called when this workstation starts the game, thus
'making it the server.  We only continue when the network is squared away.

NetworkState = NS_WAITING_TO_START  'Something other than NS_IDLE.
If NumNetworkPlayers > 0 Then
  'We do in fact have network players.  Connect with them.
  MyNetworkRole = NW_SERVER
  GameMode = GM_DIALOG_OPEN
  NetworkForm.Show vbModal, Me
  GameMode = GM_TITLE_SCREEN
End If
SetupNetwork = NetworkState

End Function

Private Sub MenuJoinGame_Click()

Dim i As Integer

MyNetworkRole = NW_CLIENT
NetworkState = NS_IDLE
GameMode = GM_DIALOG_OPEN
NetworkForm.Show vbModal, Me
GameMode = GM_TITLE_SCREEN
If NetworkState = NS_IDLE Then Exit Sub

'Other initialization goes here.
Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
SetupMenusForInGame

For i = 1 To 6
  'Stop players from making moves!
  NumOccupied(i) = -1
Next

TurnCounter = 0

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Reset stats.  No high scores for net game.
StatScreen.ResetAllStats

InitGameStuff

End Sub

Private Sub StartNetworkGame()

'At this point, our map is totally created, HQs have been picked if
'necessary, and the starting player has been determined.  Now we need
'to send *all* of this information to the other players if we're a server.
If MyNetworkRole = NW_SERVER Then
  SendStartupDataToAll
  GameMode = GM_DIALOG_OPEN
  NetworkForm.Show vbModal, Me
  GameMode = GM_GAME_ACTIVE
End If

End Sub

Private Sub MenuNew_Click()

Dim i As Long

If SetupNetwork = NS_IDLE Then Exit Sub

Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
GameMode = GM_BUILDING_MAP
SetupMenusForInGame

For i = 1 To 6
  'Stop players from making moves!
  NumOccupied(i) = -1
Next

TurnCounter = 0
Set MyMap = New Map
SetupForm

'Collect all menu settings.
CollectMapSettings (True)
CollectMenuSettings (True)
CollectDrawSettings

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Reset stats and such.
HiScores.ClearHiScores
StatScreen.ResetAllStats

'This is the syntax for building a map.
MyMap.CreateMap CfgNumCountries, MaxCountrySize, MinLakeSize, LandPct, PropPct, _
   ShapePct, CoastPctKeep, IslePctKeep

InitGameStuff

'DEBUG STUFF:  JASON, REMOVE THIS FOR RELEASE.
WriteDebugInfo

StartNetworkGame

End Sub

Private Sub MenuSame_Click()

Dim Doit As Integer
Dim Oldstr As String
Dim i As Long

If LastMapPath = "" Then Exit Sub   'No last map.

For i = 1 To 6
  'Stop players from making moves!
  NumOccupied(i) = -1
Next

Oldstr = Commentary.Caption
Commentary.Caption = "Loading map..."
DoEvents

'Load the last map, no dialog, no nothin'.
Doit = ReadMapDataKitchenSink(LastMapPath, False)

If Doit = 0 Then
  'We had an error.  Leave.
  Commentary.Caption = Oldstr
  Exit Sub
End If

If SetupNetwork = NS_IDLE Then Exit Sub

CfgNumCountries = MyMap.NumberOfCountries

TurnCounter = 0

Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
GameMode = GM_BUILDING_MAP
SetupMenusForInGame

'Collect menu settings.
CollectMenuSettings (True)
CollectDrawSettings

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Reset stats and such.
StatScreen.ResetAllStats

InitGameStuff

'DEBUG STUFF:  JASON, REMOVE THIS FOR RELEASE.
WriteDebugInfo

StartNetworkGame

End Sub

Private Sub MenuGetOptionsFromMapFile_Click()

LoadMap True

End Sub

Private Sub MenuChat_Click()

ChatForm.Show vbModeless

End Sub

Private Sub MenuLoad_Click()

Dim Doit As Integer
Dim Oldstr As String
Dim i As Long

For i = 1 To 6
  'Stop players from making moves!
  NumOccupied(i) = -1
Next

Oldstr = Commentary.Caption
Commentary.Caption = "Loading map..."
DoEvents

'This is where we actually load the map.
Doit = LoadMap(False)

If Doit = 0 Then
  'We canceled the load dialog or had an error.  Leave.
  Commentary.Caption = Oldstr
  Exit Sub
End If

If SetupNetwork = NS_IDLE Then Exit Sub

CfgNumCountries = MyMap.NumberOfCountries

TurnCounter = 0

Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
GameMode = GM_BUILDING_MAP
SetupMenusForInGame

'Collect menu settings.
CollectMenuSettings (True)
CollectDrawSettings

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Reset stats and such.
StatScreen.ResetAllStats

InitGameStuff

'DEBUG STUFF:  JASON, REMOVE THIS FOR RELEASE.
WriteDebugInfo

StartNetworkGame

End Sub

Private Sub MenuLoadSavedGame_Click()

Dim Doit As Integer
Dim Oldstr As String
Dim i As Long

For i = 1 To MAX_PLAYERS
  'Stop players from making moves!
  NumOccupied(i) = -1
Next

Oldstr = Commentary.Caption
Commentary.Caption = "Loading game..."
Commentary.Refresh
'DoEvents

'This is where we actually load the map.
Doit = LoadGame

If Doit = 0 Then
  'We canceled the load dialog or had an error.  Leave.
  Commentary.Caption = Oldstr
  Exit Sub
End If

CfgNumCountries = MyMap.NumberOfCountries

Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
GameMode = GM_BUILDING_MAP
SetupMenusForInGame

'Collect menu settings.
CollectMenuSettings (True)
CollectDrawSettings

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

'Clear the previous display.
DrawBkg

'Reset our anti-double-clickin' flags.
AlreadyAddedTroopsThisPhase = False
AlreadyPerformedActionThisPhase = False
AlreadyMovedTroopsThisPhase = False
AlreadyChoseTroopsThisPhase = False
AlreadyCheckedForMaxedOutCountries = False
AlreadyPassedThisPhase = False

'Set up the onscreen indicators.
SetupIndicators

MapMade = True
CheckedForEvent = False
EventInProgress = False

'Draw the map.
DrawMap

'Set up first message.
UpdateMessages
Me.Refresh

'DEBUG STUFF:  JASON, REMOVE THIS FOR RELEASE.
WriteDebugInfo

GameMode = GM_GAME_ACTIVE

End Sub

Private Sub InitGameStuff()

'Called when a new game is about to begin.
NetworkArbitrationInProgress = True

'Clear the previous display.
DrawBkg

'Reset all game data.
ResetGameData

'Set up initial troop positions based on menu settings.
If MyNetworkRole = NW_CLIENT Then
  CalculateTotals
Else
  PlaceTroops
End If

'Set up the onscreen indicators.
SetupIndicators
CntryName.Visible = False
AttackTxt.Visible = False
AttackTot.Visible = False
AttackLnd.Visible = False
AttackWtr.Visible = False
DefendTxt.Visible = False
DefendTot.Visible = False
DefendLnd.Visible = False
DefendWtr.Visible = False

MapMade = True
CheckedForEvent = False
EventInProgress = False
'Kick it off.
GameMode = GM_GAME_ACTIVE

'Draw the map.
DrawMap

'Set up initial HQs if we need to.
If MyNetworkRole <> NW_CLIENT Then AutoHQSelection

'Set up the initial message.
UpdateMessages
Me.Refresh

NetworkArbitrationInProgress = False

End Sub

Private Sub MenuSave_Click()

'Save the current map.
Call SaveMap

End Sub

Private Sub MenuSaveGame_Click()

'Save the current game.
If MyMap.MapName <> "" Then
  Call SaveGame(Turn, Phase)
Else
  MsgBox "You must save this map before you can save the game.", vbOKOnly Or vbExclamation, "Map Not Saved"
End If

End Sub

Private Sub MenuHiScores_Click()

'Display the Hi Scores dialog, if we're in a game...
'...or the ones for the last game.

Dim TempMode As Integer

TempMode = GameMode

GameMode = GM_DIALOG_OPEN
  HiScores.Show vbModal
GameMode = TempMode  'Restore the old mode.

End Sub

Private Sub MenuStats_Click()

Dim TempMode As Integer

'Display the Statistics dialog, if we're in a game.
If GameMode <> GM_GAME_ACTIVE Then Exit Sub

GameMode = GM_DIALOG_OPEN
  StatScreen.Show vbModeless
GameMode = GM_GAME_ACTIVE  'Restore the old mode.

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'We want to get out at any time by pressing ESC.
If KeyCode = vbKeyEscape Then
  MenuExit_Click
End If

'Shift-D brings up our debug form.
If KeyCode = vbKeyD And Shift = 1 Then
  DebugForm.Show
End If

End Sub

Private Sub BallTimer_Timer()

'This sub animates the balls.
If (GameMode <> GM_GAME_ACTIVE) And (GameMode <> GM_TITLE_SCREEN) And (GameMode <> GM_DIALOG_OPEN) Then Exit Sub

BallTimerSub  'The BallTimerSub is in the Boom.bas module.

End Sub

Private Sub AutoHQSelection()

'See if we need to quickly pick HQs for everyone (automatic HQ select).

Dim i As Integer
Dim CurrentPick As Integer
Dim RoundAbout As Integer
Dim Done As Boolean

If CFGHQSelect = 2 Then
  UpdateMessages
  Me.Refresh
  'It's the first round and we need to grab an HQ fast.
  If CFG1st = 7 Then
    CurrentPick = Int(Rnd(1) * 6) + 1
  Else
    CurrentPick = CFG1st
  End If
  RoundAbout = CurrentPick
  Done = False
  Do Until Done = True
    If PlayerType(CurrentPick) > PTYPE_INACTIVE Then Call AIchooseHQ(True, CurrentPick, 1)
    CurrentPick = CurrentPick + 1
    If CurrentPick = 7 Then CurrentPick = 1
    If CurrentPick = RoundAbout Then Done = True
  Loop
  DrawMap
  SetupIndicators
End If

End Sub

Private Sub ResetGameData()

Dim TurnOK As Boolean
Dim i As Long
Dim j As Long

'This function basically resets everything so that a new game is started.

Players.CalcNumPlayers

For i = 1 To MAX_PLAYERS
  NumOccupied(i) = 0
  NumTroops(i) = 0
  For j = 1 To MAX_PLAYERS
    'We start off liking everyone.
    Hate(i, j) = 1
  Next j
Next i

Phase = 1

If MyNetworkRole <> NW_CLIENT Then
  'Reset Country data.
  For i = 1 To MyMap.NumberOfCountries
    'Mark country as unoccupied.
    MyMap.Owner(i) = 0
    MyMap.CountryType(i) = 0
    'Assign it an 'unoccupied' color.
    MyMap.CountryColor(i) = CFGUnoccupiedColor
  Next i
  'If we want a random 1st player, make sure we pick one that isn't inactive.
  If CFG1st = 7 Then
    TurnOK = False
    Do Until TurnOK = True
      Turn = Int(Rnd(1) * 6) + 1
      If PlayerType(Turn) > PTYPE_INACTIVE Then TurnOK = True
    Loop
  Else
    'The menu specifies who's first.
    Turn = CFG1st
  End If
End If

'Reset our anti-double-clickin' flags.
AlreadyAddedTroopsThisPhase = False
AlreadyPerformedActionThisPhase = False
AlreadyMovedTroopsThisPhase = False
AlreadyChoseTroopsThisPhase = False
AlreadyCheckedForMaxedOutCountries = False
AlreadyPassedThisPhase = False

End Sub

Private Sub MenuAbout_Click()

About.Show vbModal

End Sub

Private Sub MenuExit_Click()

'Store away all of our menu settings for the next session.
MakeINI INIpath

Call UnloadAll(Me)

End Sub

Public Sub UnloadAll(Optional ByRef frmUnloadLast As Form = Nothing)
    'This will unload all the forms in the program, with the specified
    'form unloading last
    
    Dim frmFormCounter      As Form     'used to cycle through the Forms collection when unloading
    
    GameEnding = True
    
    'cycle through all the forms in the project
    For Each frmFormCounter In Forms
        
        'first make sure that this form has not been set to Nothing
        'as this is sometimes necessary to clear memory
        If Not frmFormCounter Is Nothing Then
            
            'make sure that this form is not the form that we want
            'to unload last
            If Not frmUnloadLast Is Nothing Then
                If (frmFormCounter.Name <> frmUnloadLast.Name) Then
                    Unload frmFormCounter
                End If
                
            Else
                'just unload the form - it doesn't match the one we
                'want to unload last
                Unload frmFormCounter
            End If  'is there a form to unload last
        End If  'is there a form to unload
    Next frmFormCounter
    
    'unload the last form is one was specified
    If Not frmUnloadLast Is Nothing Then
        Unload frmUnloadLast
    End If
End Sub

Private Sub MenuSfx_Click(Index As Integer)

Dim i As Long

For i = 1 To 2
  MenuSfx(i).Checked = False
Next i
MenuSfx(Index).Checked = True
CFGSound = Index

'Make this change immediately!
Call CollectDrawSettings

End Sub

Private Sub MenuSize_Click(Index As Integer)

Dim i As Long

For i = 1 To 6
  MenuSize(i).Checked = False
Next i
MenuSize(Index).Checked = True
CFGCountrySize = Index

End Sub

Private Sub MenuLakeSize_Click(Index As Integer)

Dim i As Long

For i = 1 To 4
  MenuLakeSize(i).Checked = False
Next i
MenuLakeSize(Index).Checked = True
CFGLakeSize = Index

End Sub

Private Sub MenuPct_Click(Index As Integer)

Dim i As Long

For i = 1 To 5
  MenuPct(i).Checked = False
Next i
MenuPct(Index).Checked = True
CFGLandPct = Index

End Sub

Private Sub MenuIslands_Click(Index As Integer)

Dim i As Long

For i = 1 To 3
  MenuIslands(i).Checked = False
Next i
MenuIslands(Index).Checked = True
CFGIslands = Index

End Sub

Private Sub MenuResolution_Click(Index As Integer)

Dim i As Long

'Only allowable on the title screen.
If GameMode <> GM_TITLE_SCREEN Then Exit Sub

For i = 1 To 3
  MenuResolution(i).Checked = False
Next i
MenuResolution(Index).Checked = True
CFGResolution = Index

'Update the screen!
ChangeRes

End Sub

Private Sub MenuShape_Click(Index As Integer)

Dim i As Long

For i = 1 To 3
  MenuShape(i).Checked = False
Next i
MenuShape(Index).Checked = True
CFGShape = Index

End Sub

Private Sub MenuProp_Click(Index As Integer)

Dim i As Long

For i = 1 To 3
  MenuProp(i).Checked = False
Next i
MenuProp(Index).Checked = True
CFGProportion = Index

End Sub

Private Sub MenuBorders_Click(Index As Integer)

Dim i As Long

For i = 1 To 6
  MenuBorders(i).Checked = False
Next i
MenuBorders(Index).Checked = True
CFGBorders = Index

'Changed a draw setting, let's redraw!
Call RedrawScreen

End Sub

Private Sub MenuInitTroops_Click(Index As Integer)

Dim i As Long

For i = 1 To 5
  MenuInitTroops(i).Checked = False
Next i
MenuInitTroops(Index).Checked = True
CFGInitTroopPl = Index

'Disable no bonus troops if we start with none.
If Index = 1 Then
  MenuBonus(0).Enabled = False
Else
  MenuBonus(0).Enabled = True
End If

End Sub

Private Sub MenuInitTroopCts_Click(Index As Integer)

Dim i As Long

For i = 1 To 3
  MenuInitTroopCts(i).Checked = False
Next i
MenuInitTroopCts(Index).Checked = True
CFGInitTroopCt = Index

End Sub

Private Sub MenuShips_Click(Index As Integer)

Dim i As Long

For i = 1 To 4
  MenuShips(i).Checked = False
Next i
MenuShips(Index).Checked = True
CFGShips = Index

End Sub

Private Sub MenuPorts_Click(Index As Integer)

Dim i As Long

For i = 1 To 4
  MenuPorts(i).Checked = False
Next i
MenuPorts(Index).Checked = True
CFGPorts = Index

End Sub

Private Sub MenuConquer_Click(Index As Integer)

Dim i As Long

For i = 1 To 5
  MenuConquer(i).Checked = False
Next i
MenuConquer(Index).Checked = True
CFGConquer = Index

End Sub

Private Sub MenuBonus_Click(Index As Integer)

Dim i As Long

For i = 0 To 5
  MenuBonus(i).Checked = False
Next i
MenuBonus(Index).Checked = True
CFGBonusTroops = Index

'Disable no starting troops if we have no bonus troops.
If Index = 0 Then
  MenuInitTroops(1).Enabled = False
Else
  MenuInitTroops(1).Enabled = True
End If

End Sub

Private Sub MenuRandom_Click(Index As Integer)

Dim i As Long

For i = 1 To 3
  MenuRandom(i).Checked = False
Next i
MenuRandom(Index).Checked = True
CFGEvents = Index

End Sub

Private Sub MenuHQSelect_Click(Index As Integer)

Dim i As Long

For i = 1 To 2
  MenuHQSelect(i).Checked = False
Next i
MenuHQSelect(Index).Checked = True
CFGHQSelect = Index

End Sub

Private Sub Menu1st_Click(Index As Integer)

Dim i As Long

For i = 1 To 7
  Menu1st(i).Checked = False
Next i
Menu1st(Index).Checked = True
CFG1st = Index

End Sub

Private Sub MenuAISpeed_Click(Index As Integer)

Dim i As Long

For i = 1 To 4
  MenuAISpeed(i).Checked = False
Next i
MenuAISpeed(Index).Checked = True
CFGAISpeed = Index

End Sub

Private Sub MenuRefresh_click()

'Let's redraw!
If GameMode = GM_GAME_ACTIVE Or GameMode = GM_TITLE_SCREEN Then
  RedrawScreen
  UpdateMessages
End If

End Sub

Private Sub MenuTextFile_click()

If Dir(HLPpath) <> "" Then
  Shell "winhlp32.exe " & HLPpath, vbNormalFocus
Else
  MsgBox "Could not find Fracas.hlp.", vbOKOnly, "Fracas Help"
End If

End Sub

Private Sub MenuPlayers_click()

Dim TempMode As Integer

TempMode = GameMode

GameMode = GM_DIALOG_OPEN
  Players.Show vbModeless
GameMode = TempMode  'Restore the old mode.

End Sub

Public Sub RedrawScreen()

Dim i As Long

'Grab the display settings.
Call CollectDrawSettings

If GameMode = GM_GAME_ACTIVE Then
  'Set up the CountryColor array.
  For i = 1 To MyMap.NumberOfCountries
    If MyMap.Owner(i) = 0 Then
      MyMap.CountryColor(i) = CFGUnoccupiedColor
    End If
  Next i
End If

'Clear the previous display.
Call DrawBkg

'Draw the map!
Call DrawMap

End Sub

Public Sub CollectDrawSettings()

'This subroutine grabs the display info from the menus.

Dim i As Long

For i = 1 To MAX_MENU_ITEMS
  If i <= 6 Then
    If MenuBorders(i).Checked = True Then CFGBorders = i
  End If
  If i <= 4 Then
    If MenuAISpeed(i).Checked = True Then CFGAISpeed = i
  End If
  If i <= 2 Then
    If MenuSfx(i).Checked = True Then CFGSound = i
  End If
Next i

End Sub

Public Sub CollectMenuSettings(Parse As Boolean)

'This subroutine grabs the game info from the menus.

Dim i As Long

For i = 0 To MAX_MENU_ITEMS
  If i <= 7 And i > 0 Then
    If Menu1st(i).Checked = True Then CFG1st = i
  End If
  If i <= 5 Then
    If MenuBonus(i).Checked = True Then CFGBonusTroops = i
  End If
  If i <= 5 And i > 0 Then
    If MenuInitTroops(i).Checked = True Then CFGInitTroopPl = i
    If MenuConquer(i).Checked = True Then CFGConquer = i
  End If
  If i <= 4 And i > 0 Then
    If MenuShips(i).Checked = True Then CFGShips = i
    If MenuPorts(i).Checked = True Then CFGPorts = i
  End If
  If i <= 3 And i > 0 Then
    If MenuRandom(i).Checked = True Then CFGEvents = i
    If MenuInitTroopCts(i).Checked = True Then CFGInitTroopCt = i
  End If
  If i <= 2 And i > 0 Then
    If MenuHQSelect(i).Checked = True Then CFGHQSelect = i
  End If
Next i

'Set the number of players here, since this function is called
'every time we start a new game.
NumPlayers = Players.CalcNumPlayers

'Now parse the menu settings.
If Parse = False Then Exit Sub

'Set up the overwater offensive/defensive percentage.
Select Case CFGShips
  Case 1:    ShipPct = 0.1
  Case 2:    ShipPct = 0.25
  Case 3:    ShipPct = 0.5
  Case 4:    ShipPct = 1
End Select

End Sub

Public Sub CollectMapSettings(Parse As Boolean)

'This subroutine grabs the map config info from the menus.

Dim i As Long
Dim SizeDiv As Long

For i = 0 To MAX_MENU_ITEMS
  If i <= 6 And i > 0 Then
    If MenuSize(i).Checked = True Then CFGCountrySize = i
  End If
  If i <= 5 And i > 0 Then
    If MenuPct(i).Checked = True Then CFGLandPct = i
  End If
  If i <= 4 And i > 0 Then
    If MenuLakeSize(i).Checked = True Then CFGLakeSize = i
  End If
  If i <= 3 And i > 0 Then
    If MenuIslands(i).Checked = True Then CFGIslands = i
    If MenuResolution(i).Checked = True Then CFGResolution = i
    If MenuShape(i).Checked = True Then CFGShape = i
    If MenuProp(i).Checked = True Then CFGProportion = i
  End If
Next i

'Now parse the menu settings.
If Parse = False Then Exit Sub

'Get the number of countries that will fit in the
'selected area based on country size.
If CFGCountrySize < 6 Then
  'Stair-step what our size should be.
  SizeDiv = Int((MAX_COUNTRY_SIZE - MIN_COUNTRY_SIZE) / 4)
  MaxCountrySize = (SizeDiv * CFGCountrySize) + (MIN_COUNTRY_SIZE - SizeDiv)
Else
  'Hodge-Podge.  Average the value for now.
  MaxCountrySize = (MAX_COUNTRY_SIZE + MIN_COUNTRY_SIZE) / 2
End If

Select Case CFGLandPct:
  Case 1:    LandPct = 0.1
  Case 2:    LandPct = 0.25
  Case 3:    LandPct = 0.5
  Case 4:    LandPct = 0.75
  Case 5:    LandPct = 0.9
End Select
CfgNumCountries = Int((MyMap.Xsize * MyMap.Ysize * LandPct) / MaxCountrySize)
If CfgNumCountries = 0 Then CfgNumCountries = 1
If CfgNumCountries >= 999 Then CfgNumCountries = 998
'Make sure the number of countries is evenly divisible by the number of players.
If NumPlayers = 0 Then Exit Sub    'Prevent /0.
If Int(CfgNumCountries / NumPlayers) <> CfgNumCountries / NumPlayers Then
  CfgNumCountries = CfgNumCountries + (NumPlayers - (CfgNumCountries Mod NumPlayers))
End If

'Get the proportional size variance.
Select Case CFGProportion:
  Case 1:    PropPct = MAX_PROP_PCT / 4
  Case 2:    PropPct = MAX_PROP_PCT / 2
  Case 3:    PropPct = MAX_PROP_PCT
End Select

'Get the minimum allowable lake size.
Select Case CFGLakeSize
  Case 1:    MinLakeSize = 0
  Case 2:    MinLakeSize = 5
  Case 3:    MinLakeSize = 10
  Case 4:    MinLakeSize = 20
End Select

'Get the percentage irregularity.
Select Case CFGShape
  Case 1:    ShapePct = 1
  Case 2:    ShapePct = 0.5
  Case 3:    ShapePct = 0.25
End Select

'Get the island parameters.
If CFGIslands = 2 Then
  CoastPctKeep = 0.99
  IslePctKeep = 0.01
End If

End Sub

Public Sub DrawBkg()

'BitBlt MapBuffer.hdc, 0, 0, Wide, Tall, Ocean.hdc, 0, 0, SRCCOPY

Dim i As Long
Dim j As Long

'Tile the ocean many times in each direction.
For i = 0 To Wide Step OCEAN_X_SIZE
  For j = 0 To Tall Step OCEAN_Y_SIZE
    BitBlt MapBuffer.hdc, i, j, OCEAN_X_SIZE, OCEAN_Y_SIZE, Ocean.hdc, 0, 0, SRCCOPY
  Next j
Next i

End Sub

Public Sub CreateOcean()

Dim i As Long
Dim j As Long

'Build a chunk of the ocean.
For i = 0 To OCEAN_X_SIZE
  For j = 0 To OCEAN_Y_SIZE
    Call SetPixel(PicBuffer.hdc, i, j, WaterColors(GetWaterColor))
  Next j
Next i

'Tile it many times in each direction.
For i = 0 To Wide Step OCEAN_X_SIZE
  For j = 0 To Tall Step OCEAN_Y_SIZE
    Call BitBlt(Ocean.hdc, i, j, OCEAN_X_SIZE, OCEAN_Y_SIZE, PicBuffer.hdc, 0, 0, SRCCOPY)
  Next j
Next i

End Sub

Public Function GetWaterColor() As Integer

  Dim r As Single
  
  r = Rnd(1)

  If r < 0.1 Then
    GetWaterColor = 1
  ElseIf r < 0.2 Then
    GetWaterColor = 2
  ElseIf r < 0.3 Then
    GetWaterColor = 3
  ElseIf r < 0.4 Then
    GetWaterColor = 4
  ElseIf r < 0.9 Then
    GetWaterColor = 5
  Else
    GetWaterColor = 6
  End If

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim MapX As Integer
Dim MapY As Integer
Dim TempNumb1 As Integer
Dim TempNumb2 As Integer
Dim i As Long
Dim j As Long
Dim k As Integer
Dim CurrentMouse As Long

If GameMode <> GM_GAME_ACTIVE Then Exit Sub
If EventInProgress = True Then Exit Sub

'Calculate the map coordinates where the mouse is.
MapX = Int((Int(x / Screen.TwipsPerPixelX) - XOFFSET) / 8) + 1
MapY = Int((Int(y / Screen.TwipsPerPixelY) - YOFFSET) / 8) + 1
'Identify the country we're over.
If (MapX > 0) And (MapX <= MyMap.Xsize) And (MapY > 0) And (MapY <= MyMap.Ysize) Then
  CurrentMouse = MyMap.Grid(MapX, MapY)
Else
  'We shouldn't be clicking outside the map *anyway*.
  Exit Sub
End If
  
  If (MapX > 0) And (MapX <= MyMap.Xsize) And (MapY > 0) And (MapY <= MyMap.Ysize) Then
    If Button = 1 Then
      'The user clicked the left mouse button.  If it is the user's turn,
      'we go to the turn sub.  If it is a computer player's turn, exit.

      'Start a small splash if we're clicking in water.
      If CurrentMouse > TILEVAL_COASTLINE Then
        Call BuildBalls(SPLASH_COUNT, _
                    Int(x / Screen.TwipsPerPixelX) - XOFFSET, _
                    Int(y / Screen.TwipsPerPixelY) - YOFFSET, _
                    SPLASH_INTENSITY, _
                    SPLASH_SPREAD, _
                    SPLASH_ELASTIC, _
                    SPLASH_SIZE, _
                    SPLASH_COLOR)
        Call SNDSplishSplash
        Exit Sub
      End If

      If PlayerType(Turn) <> PTYPE_HUMAN Then Exit Sub
      
      Call PlayerClick(CurrentMouse, Turn, Phase)   'In the Play.bas module.
      
    ElseIf Button = 2 Then
      'what happens when we click the right mouse button.
      If CurrentMouse >= TILEVAL_COASTLINE Then
        'Make a small splash, but different than the left water-click.
        Call BuildBalls(SPLASH_COUNT + 10, _
                    Int(x / Screen.TwipsPerPixelX) - XOFFSET, _
                    Int(y / Screen.TwipsPerPixelY) - YOFFSET, _
                    SPLASH_INTENSITY + 3, _
                    SPLASH_SPREAD, _
                    SPLASH_ELASTIC, _
                    SPLASH_SIZE, _
                    SPLASH_COLOR)
        Call PopulateCountryFrame(Turn, CurrentMouse)
        Call SNDSplishSplash
        LastRightClick = CurrentMouse
        Exit Sub
      End If
      'If we right-clicked on land, we need to populate the items on the right side.
      'First, see how many human players there are.
      TempNumb1 = 0
      TempNumb2 = 0
      For i = 1 To 6
        If PlayerType(i) = PTYPE_HUMAN Then
          TempNumb1 = TempNumb1 + 1
          TempNumb2 = i
        End If
      Next i
      If TempNumb1 = 1 Then
        'Only one human playing, so populate the table with their stats.
        Call CalculateStrengths(TempNumb2, CurrentMouse)
        Call PopulateCountryFrame(TempNumb2, CurrentMouse)
      Else
        'More than one human.  If it's a human's turn, do like normal.
        If PlayerType(Turn) = PTYPE_HUMAN Then
          Call CalculateStrengths(Turn, CurrentMouse)
          Call PopulateCountryFrame(Turn, CurrentMouse)
        Else
          'We're right-clicking during a computer's turn!  Put up JUST
          'defense stats.
          Call CalculateStrengths(Turn, CurrentMouse)
          Call PopulateCountryFrame(Turn, CurrentMouse)
          AttackTxt.Visible = False
          AttackTot.Visible = False
          AttackLnd.Visible = False
          AttackWtr.Visible = False
        End If
      End If
      'Now flash either the right-clicked country or everything it borders.
      'Basically, the second right-click on a country will show its neighbors.
      If LastRightClick = CurrentMouse Then
        'Flash our neighbors.  If we have a port, do the overseas ones too.
        For j = 1 To MyMap.NumberOfCountries
          TempNumb1 = AIxCanReachy(CurrentMouse, j)
          If (TempNumb1 = 1) Or ((TempNumb1 = 2) And _
                              ((MyMap.CountryType(CurrentMouse) And 1) = 1)) Then
            Call LongerFlash(j)
          End If
        Next j
        'Also erase the flash on the country itself (it looks cooler).
        For k = 1 To 4
          MyMap.Flash(CurrentMouse, k) = 0
        Next k
        'Update everything and get it flashing the first time.
        DrawMap
        LastRightClick = -LastRightClick  'Flag it as clicked twice so renames still work.
      Else
        'Flash just us.
        Call LongFlash(CurrentMouse)
        DrawMap
        LastRightClick = CurrentMouse
      End If
    End If
  End If

End Sub

Public Function Explode(CountryID As Long, ExpColor As Integer)

'Start a big ass explosion based on country size and troop count.
If CountryID < TILEVAL_COASTLINE And CountryID > 0 Then
  If MyMap.TroopCount(CountryID) > 0 Then
    Select Case (CFGCountrySize)
    Case 1:
      Call BuildBalls(12 + (MyMap.TroopCount(CountryID) * SMALL_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    15 + SMALL_INTENSITY * MyMap.TroopCount(CountryID), _
                    SMALL_SPREAD, _
                    SMALL_ELASTIC, _
                    SMALL_SIZE, _
                    ExpColor)
    Case 2:
      Call BuildBalls(15 + (MyMap.TroopCount(CountryID) * SMALL_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    16 + SMALL_INTENSITY * MyMap.TroopCount(CountryID), _
                    SMALL_SPREAD + 2, _
                    SMALL_ELASTIC, _
                    SMALL_SIZE + 1, _
                    ExpColor)
    Case 3:
      Call BuildBalls(18 + (MyMap.TroopCount(CountryID) * MED_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    17 + MED_INTENSITY * MyMap.TroopCount(CountryID), _
                    MED_SPREAD, _
                    MED_ELASTIC, _
                    MED_SIZE, _
                    ExpColor)
    Case 4:
      Call BuildBalls(21 + (MyMap.TroopCount(CountryID) * MED_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    18 + MED_INTENSITY * MyMap.TroopCount(CountryID), _
                    MED_SPREAD + 2, _
                    MED_ELASTIC, _
                    MED_SIZE + 1, _
                    ExpColor)
    Case 5, 6:
      Call BuildBalls(24 + (MyMap.TroopCount(CountryID) * BIG_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    20 + BIG_INTENSITY * MyMap.TroopCount(CountryID), _
                    BIG_SPREAD, _
                    BIG_ELASTIC, _
                    BIG_SIZE, _
                    ExpColor)
    End Select
  Else
    'This country just got taken over.  Let's make a big balls explosion for it.
     Call BuildBalls(1 + (250 * BIG_COUNT), _
                    (MyMap.DigitCoords(CountryID, 1) - 1) * 8, _
                    (MyMap.DigitCoords(CountryID, 2) - 1) * 8, _
                    35 + (BIG_INTENSITY * 200), _
                    BIG_SPREAD + 5, _
                    BIG_ELASTIC, _
                    BIG_SIZE, _
                    ExpColor)

  End If
End If

End Function

Private Sub GameTimer_Timer()

Dim WinPl As Integer
Dim TitX As Long
Dim TitY As Long
Dim i As Long
Dim TempBool As Boolean

'This timer controls the game sequencing.  If it is a human's turn,
'then nothing happens here.  But if it is a computer's turn, we run
'through the AI subroutines at the appropriate turn phases.

'Draw the title map if it's not there already.
If GameMode = GM_BUILDING_TITLE Then
  If MapMade = True Then
    CfgNumCountries = 4
  Else
    'This is the first time the map is being created.
    Set MyMap = New Map
  End If
  
  'DoEvents
  
  'If the main form is minimized, get it back up!
  If Land.WindowState = vbMinimized Then
    Land.WindowState = vbNormal
  End If
 
  SetupForm
  'DoEvents
  MyMap.SetupTitleScreen
  Call DrawBkg
  
  'DoEvents
  Call DrawMap
  Me.Refresh
  
  GameMode = GM_TITLE_SCREEN  'Title is now built.
End If
  
'Is there currently no game in progress?
'If so, we need to do titlescreen activities.
If GameMode = GM_TITLE_SCREEN Then
  'Then we do the cadence once every 8 seconds.
  If Time >= CadenceTime + "00:00:08" Then
    CadenceTime = Time
    Call SNDPlayWarCadence
    'Put a random explosion on the title screen (land only) every once in a while.
    Do
      TitX = Int(Rnd(1) * MyMap.Xsize) + 1
      TitY = Int(Rnd(1) * MyMap.Ysize) + 1
    Loop Until MyMap.Grid(CInt(TitX), CInt(TitY)) < TILEVAL_COASTLINE
    
    Call BuildBalls(70 + Int(Rnd(1) * 100) + 1, _
                    (TitX - 1) * 8, _
                    (TitY - 1) * 8, _
                    30 + Int(Rnd(1) * 40) + 1, _
                    MED_SPREAD + Int(Rnd(1) * 30) + 1, _
                    MED_ELASTIC, _
                    BIG_SIZE, _
                    Int(Rnd(1) * 11) + 1)
  End If
End If

'Only worry about this junk if a game is underway.
If (GameMode <> GM_GAME_ACTIVE) Or (NetworkArbitrationInProgress = True) Then Exit Sub

'First, if this player is inactive or already dead, move up to the next player.
If PlayerType(Turn) = PTYPE_INACTIVE Or NumOccupied(Turn) = -1 Then
  Call NextTurn(Turn, Phase)
  Exit Sub
End If

'Do any outstanding events...
If EventInProgress = True Then
  Call RandomEventDoer
  Exit Sub
End If

'Let's check to see if there is only one player left.
'If so, then this is the winner!
'TODO:  Update this to include servers and clients.
WonTurn = 0
For i = 1 To 6
  If (PlayerType(i) > PTYPE_INACTIVE) And (NumOccupied(i) > -1) Then
    WonTurn = WonTurn + 1
    WinPl = i
  End If
Next i

If WonTurn = 1 Then
'Only one person left alive!
  'Get rid of the last flash.
  ResetFlashing
  DrawMap
  'Stop the game and all timers.
  WonGame = False
  Do Until WonGame = True
    WonGame = NoBalls      'This function is in Boom.bas.
    DoEvents               'we need this or the program will get stuck in an infinite loop
  Loop
  
  DoEvents
  
  Commentary.Caption = PlayerName(WinPl) & " has won the game!"
  Commentary.BackColor = Land.BackColor
  Commentary.ForeColor = vbBlack
  GameMode = GM_DIALOG_OPEN
  Tribute.Show vbModeless
  Exit Sub
End If

'Get out if all players haven't picked HQ yet during auto HQ selection.
If PickedHQYet = False Then Exit Sub

'If we're in the reinforcement phase, check some things.
If (Phase = 1) And (NumOccupied(Turn) > 0) Then
  'Check if we have no reinforcements.  If not, skip this phase.
  If CFGBonusTroops = 0 Then
    Call NextPhase(Turn, Phase)
    Exit Sub
  Else
    'Check if we have nowhere to put troops!  If everything we have is
    'maxed out, skip this phase.  Bugfix for version 2.0 BETA.
    If AlreadyCheckedForMaxedOutCountries = False Then
      For i = 1 To MyMap.NumberOfCountries
        'If this country belongs to the current player and is not maxed out...
        If (MyMap.Owner(i) = Turn) And (MyMap.TroopCount(i) < MAX_COUNTRY_CAPACITY) Then
          '...then stop looking, we *can* place troops this turn.
          AlreadyCheckedForMaxedOutCountries = True
          Exit For
        End If
      Next i
      'If we found just one country that's not maxed, keep going.  Otherwise leave.
      If AlreadyCheckedForMaxedOutCountries = False Then
        AlreadyCheckedForMaxedOutCountries = True
        Call NextPhase(Turn, Phase)
        Exit Sub
      End If
    End If
  End If
End If

'If this is a human, we exit and basically wait for them to finish.
'Or if this is a networked player, we wait on their input.
If (PlayerType(Turn) = PTYPE_HUMAN) Or (PlayerType(Turn) = PTYPE_NETWORK) Or _
   (PlayerType(Turn) = PTYPE_SERVER) Then Exit Sub

'If we got here, then this is a computer!

'Now see what we'll do.
If NumOccupied(Turn) = 0 Then
  'Timing fix: If we got here, and auto HQ selection is enabled, leave!
  If CFGHQSelect = 2 Then Exit Sub   'We shouldn't be picking our own HQ.
  'This is their first turn.
  Call AIdelay
  Call AIchooseHQ(False, Turn, Phase)
  Call NextTurn(Turn, Phase)
  Exit Sub
End If
  
Select Case Phase
Case 1:
  'Add reinforcements.
  Call AIdelay
  Call AIreinforce(Turn, Phase)
  Call NextPhase(Turn, Phase)
  Exit Sub
Case 2:
  'Action!
  Call AIdelay
  Call AIaction(Turn, Phase)
  Call NextPhase(Turn, Phase)
  Exit Sub
Case 3:
  'Troop movement.
  Call AIdelay
  Call AItroopmove(Turn, Phase)
  Call NextTurn(Turn, Phase)
  MsgXferAmount = 0  'Reset the total in the troop move message.
  Exit Sub
End Select

End Sub

Private Sub FlashTimer_Timer()

'Let's do a cycle of the Flashing engine real quick.
If GameMode = GM_GAME_ACTIVE Then FlashTimerSub

End Sub

Public Sub UpdateMessages()

'This sub puts the appropriate message in the commentary line.

Dim i As Long
Dim j As Integer

'Leave if we don't need messaging.
If (GameMode <> GM_GAME_ACTIVE) Then Exit Sub

'Get out of here if we don't need to print anything.
If (NumOccupied(Turn) = -1) Or (PlayerType(Turn) = PTYPE_INACTIVE) Then Exit Sub

'Fix the Resign button's visibility.
If (PlayerType(Turn) = PTYPE_HUMAN) And (NumOccupied(Turn) > 0) And (Phase = 1) Then
  'If we're the last person left, *don't* put up the resign button.
  j = 0
  For i = 1 To MAX_PLAYERS
    If NumOccupied(i) > -1 Then j = j + 1
  Next i
  If j <> 1 Then
    ResignBut.Visible = True
  Else
    ResignBut.Visible = False
  End If
Else
  ResignBut.Visible = False
End If

'Punch in the current player's name so we all know whose turn it is.
For i = 1 To MAX_PLAYERS
  If Turn = i Then
    PlyrNames(i).BorderStyle = 1
  Else
    PlyrNames(i).BorderStyle = 0
  End If
Next i

'If we're auto-picking HQs, put up a message to that effect...
If PickedHQYet = False Then
  Commentary.BackColor = Land.BackColor
  Commentary.ForeColor = vbBlack
  Commentary.Caption = "Selecting HQs..."
  Exit Sub
End If

'Put the right color on the message text.
Commentary.BackColor = PlayerColorCodes(Player(Turn))
Commentary.ForeColor = PlayerTextColor(Player(Turn))

Select Case PlayerType(Turn)
Case PTYPE_HUMAN:
  'This is a human's turn.
  'See if they have any countries at all.
  If (NumOccupied(Turn) = 0) And (CFGHQSelect = 1) Then
    'Nope, then this is their FIRST turn.
    Commentary.Caption = " " & PlayerName(Turn) & ": Choose your first country.  This country will be your HQ for the rest of the game."
    Exit Sub
  End If
  
  Select Case Phase
  Case 1:
    'Reinforcement phase.
    If CFGBonusTroops * NumOccupied(Turn) = 1 Then
      Commentary.Caption = " " & PlayerName(Turn) & ": Place " & (CFGBonusTroops * NumOccupied(Turn)) & " little troop in one of your countries."
    Else
      Commentary.Caption = " " & PlayerName(Turn) & ": Place " & (CFGBonusTroops * NumOccupied(Turn)) & " troops in one of your countries."
    End If
    
    Exit Sub
  Case 2:
    'Action phase.
    Commentary.Caption = " " & PlayerName(Turn) & ": Choose an enemy country to attack, a neutral country to annex, or your country to build a port."
    Exit Sub
  Case 3:
    'Troop movement source select phase.
    Commentary.Caption = " " & PlayerName(Turn) & ": Choose a country to move troops from, or Pass for no troop movement."
    Exit Sub
  Case 4:
    'Troop movement destination select phase.
    Commentary.Caption = " " & PlayerName(Turn) & ": Enter the number of troops to move and choose a destination, or Pass for no troop movement."
    Exit Sub
  End Select
  
Case PTYPE_COMPUTER:
  'This is a computer's turn!
  'See if they have any countries at all.
  If (NumOccupied(Turn) = 0) And (CFGHQSelect = 1) Then
    'Nope, then this is their FIRST turn.
    Commentary.Caption = " " & PlayerName(Turn) & " is setting up headquarters..."
    Exit Sub
  End If
  
  Select Case Phase
  Case 1:
    'Reinforcement phase.
    Select Case (CFGBonusTroops * NumOccupied(Turn))
      Case 0:
        Commentary.Caption = ""
      Case 1:
        Commentary.Caption = " " & PlayerName(Turn) & " is placing one little troop..."
      Case Else
        Commentary.Caption = " " & PlayerName(Turn) & " is placing " & (CFGBonusTroops * NumOccupied(Turn)) & " troops..."
    End Select
    Exit Sub
  Case 2:
    'Action phase.
    Commentary.Caption = " " & PlayerName(Turn) & " is thinking..."
    Exit Sub
  Case 3:
    'Troop movement source select phase.
    Select Case MsgXferAmount
      Case 0:
        Commentary.Caption = " " & PlayerName(Turn) & " is thinking about a troop movement..."
      Case 1:
        Commentary.Caption = " " & PlayerName(Turn) & " is moving one little troop..."
      Case Else:
        Commentary.Caption = " " & PlayerName(Turn) & " is moving " & MsgXferAmount & " troops..."
    End Select
    Exit Sub
  Case 4:
    'Troop movement destination select phase.  We should never, ever get here.  :)
    Commentary.Caption = " " & PlayerName(Turn) & " has found a BUG!  How did you get here?"
    Exit Sub
  End Select
  
Case PTYPE_NETWORK, PTYPE_SERVER:
  'Either this is a server waiting on a client or this is a client waiting for the server.
  If (NumOccupied(Turn) = 0) And (CFGHQSelect = 1) Then
    'Nope, then this is their FIRST turn.
    Commentary.Caption = " Waiting for " & PlayerName(Turn) & " to set up headquarters..."
    Exit Sub
  End If
  
  Select Case Phase
  Case 1:
    'Reinforcement phase.
    Select Case (CFGBonusTroops * NumOccupied(Turn))
      Case 0:
        Commentary.Caption = ""
      Case 1:
        Commentary.Caption = " Waiting for " & PlayerName(Turn) & " to place one little troop..."
      Case Else
        Commentary.Caption = " Waiting for " & PlayerName(Turn) & " to place " & (CFGBonusTroops * NumOccupied(Turn)) & " troops..."
    End Select
    Exit Sub
  Case 2:
    'Action phase.
    Commentary.Caption = " Waiting for " & PlayerName(Turn) & " to complete an action..."
    Exit Sub
  Case 3:
    'Troop movement source select phase.
    Commentary.Caption = " Waiting for " & PlayerName(Turn) & " to move troops..."
    Exit Sub
  Case 4:
    'Troop movement destination select phase.  We should never, ever get here.  :)
    Commentary.Caption = " " & PlayerName(Turn) & " has found a BUG!  How did you get here?"
    Exit Sub
  End Select
  
End Select

Commentary.Caption = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Store away all of our menu settings for the next session.
MakeINI INIpath

Call UnloadAll(Me)
End
End Sub

Public Sub Form_Paint()

Dim Xbase As Integer
Const Ybase = 38.5
Const Dist = 470 / 15
Dim i As Integer
Dim IconType As Integer

'Copy the map to the screen each time the form is painted.
'This will automatically refresh the map if something eclipses it
'(like the tribute dialog).
u% = BitBlt(hdc, XOFFSET, YOFFSET, Wide, Tall, PicBuffer.hdc, 0, 0, SRCCOPY)

If GameMode <> GM_GAME_ACTIVE Then Exit Sub

Xbase = (Land.Width / Screen.TwipsPerPixelX) - 100

For i = 1 To MAX_PLAYERS
  If (MyNetworkRole = NW_SERVER) Or (MyNetworkRole = NW_NONE) Then
    Select Case PlayerType(i)
      Case PTYPE_HUMAN:  IconType = 1
      Case PTYPE_COMPUTER:  IconType = 2
      Case PTYPE_NETWORK:  IconType = 4
      Case Else:  IconType = 0
    End Select
  ElseIf MyNetworkRole = NW_CLIENT Then
    Select Case TempPlayerType(i)
      Case PTYPE_HUMAN:  IconType = 3   'Host machine human.
      Case PTYPE_COMPUTER:  IconType = 2
      Case PTYPE_NETWORK:
        If i = MyClientIndex Then
          IconType = 1   'Human on this network client.
        Else
          IconType = 4   'Human on another network client.
        End If
      Case Else:  IconType = 0
    End Select
  End If
  If IconType > 0 Then
    u% = BitBlt(hdc, Xbase, Ybase + ((i - 1) * Dist), GFX_GRID, GFX_GRID, Land!LandMap.hdc, (GFX_ICONS_X + IconType) * GFX_GRID, (GFX_ICONS_Y + 1) * GFX_GRID, SRCAND)
    u% = BitBlt(hdc, Xbase, Ybase + ((i - 1) * Dist), GFX_GRID, GFX_GRID, Land!LandMap.hdc, (GFX_ICONS_X + IconType) * GFX_GRID, GFX_ICONS_Y * GFX_GRID, SRCINVERT)
  End If
Next i

End Sub

Private Sub PlaceTroops()

Dim TMax As Long
Dim TMin As Long
Dim TRange As Long
Dim Cntry As Long
Dim Bool1 As Boolean
Dim Bool2 As Boolean
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim r As Single

'Clean out all countries first.
For i = 1 To MyMap.NumberOfCountries
  MyMap.TroopCount(i) = 0
Next i

'Leave if there are no initial troops.
If CFGInitTroopPl = 1 Then Exit Sub

Select Case CFGInitTroopCt
  Case 1:
    TMax = 10
    TMin = 1
    TRange = 10
  Case 2:
    TMax = 50
    TMin = 10
    TRange = 40
  Case 3:
    TMax = 100
    TMin = 40
    TRange = 60
End Select

'First, we will place a number of large quantities of troops around the map.
'One bunch per player, and no two next to each other.
For i = 1 To NumPlayers
  Bool1 = False
  k = 0
  Do Until Bool1 = True Or k = 100
    k = k + 1
    j = Int(Rnd(1) * MyMap.NumberOfCountries) + 1
    'Let's check all neighbors of this country for existing troops.
    If MyMap.TroopCount(j) > 0 Then
      'This one is already stacked.
      Bool1 = False
    Else
      Bool2 = False
      For l = 1 To MyMap.MaxNeighbors
        Cntry = MyMap.Neighbors(j, CInt(l))
        If (Cntry = 0) Or (Cntry >= TILEVAL_COASTLINE) Then Exit For
        If MyMap.TroopCount(Cntry) > 0 Then
          'We're next to one already.
          Bool2 = True
          Exit For
        End If
      Next l
      If Bool2 = False Then
        'We found a good one!
        Bool1 = True
        MyMap.TroopCount(j) = TMax - Int(Rnd(1) * ((TRange * 0.1)))
      End If
    End If
  Loop
  If k = 100 Then
    'We must be playing on a map with not many countries.
    'Put it in the first one we find.
    For l = 1 To MyMap.NumberOfCountries
      If MyMap.TroopCount(l) = 0 Then
        MyMap.TroopCount(l) = TMax - Int(Rnd(1) * ((TRange * 0.1)))
      End If
    Next
  End If
Next i

'Now we will populate the rest of the map with lesser quantities,
'based on the configuration settings.
For i = 1 To MyMap.NumberOfCountries
  r = Rnd(1)
  If r < (CFGInitTroopPl - 1) * 0.25 Then
    If MyMap.TroopCount(i) = 0 Then
      'This is an empty country.
      MyMap.TroopCount(i) = Int(Rnd(1) * TRange) + TMin
    End If
  End If
Next i

End Sub

Private Sub PassTurnBut_Click()

'If this button got clicked, then it means the human player has
'chosen to do nothing this phase.  We will simply bounce to the
'next phase.

If AlreadyPassedThisPhase = True Then Exit Sub

AlreadyPassedThisPhase = True

Call SNDWhistle
Call SendPassToNetwork(Turn, Phase, 0)

If Phase < 3 Then
  Call NextPhase(Turn, Phase)
Else
  'If we're somewhere in troop movement, just skip it all.
  Call NextTurn(Turn, Phase)
End If

End Sub

Public Sub SetupIndicators()

Dim i As Long

'In between turns or phases, we update the player frame.
For i = 1 To 6
  PlyrNames(i) = PlayerName(i)
  PlyrNames(i).BackColor = PlayerColorCodes(Player(i))
  PlyrNames(i).ForeColor = PlayerTextColor(Player(i))
  PlyrTotals(i) = Trim(Str(NumOccupied(i))) & " : " & Trim(Str(NumTroops(i)))
  'If a player is inactive, don't display anything.
  If PlayerType(i) = PTYPE_INACTIVE Then
    PlyrNames(i).Visible = False
    PlyrTotals(i).Visible = False
  Else
    PlyrNames(i).Visible = True
    PlyrTotals(i).Visible = True
  End If
  'If a player got killed, then gray out their name.
  If NumOccupied(i) < 0 Then
    PlyrTotals(i).Visible = False
    PlyrNames(i).Enabled = False
  ElseIf PlayerType(i) <> PTYPE_INACTIVE Then
    PlyrTotals(i).Visible = True
    PlyrNames(i).Enabled = True
  End If
Next i

'In between turns or phases, we clear out the contents of the country frame.
'Only a right-click can populate this frame!

'If the last turn was a computer, let's leave the side dialog up since a human
'was probably making some observations.  If a human just played, let's kill
'the country frame so that the next player doesn't see too much!
If PlayerType(Turn) = PTYPE_HUMAN Then
  CntryName.Visible = False
  AttackTxt.Visible = False
  AttackTot.Visible = False
  AttackLnd.Visible = False
  AttackWtr.Visible = False
  DefendTxt.Visible = False
  DefendTot.Visible = False
  DefendLnd.Visible = False
  DefendWtr.Visible = False
End If

'Show both frames.
PlyrFrame.Visible = True
CntryFrame.Visible = True

End Sub

Public Function PopulateCountryFrame(RightClicker As Integer, CurrentMouse As Long)

'This function gets called when a country is right-clicked.  It basically just
'populates the right-hand information area.

If CurrentMouse >= TILEVAL_COASTLINE Then
  'We right-clicked in water.  Let's put up the name of the water mass.
  CntryName.Caption = MyMap.WaterName(CurrentMouse - 1000)
  CntryName.BackColor = Land.BackColor
  CntryName.ForeColor = vbBlack
  CntryName.Visible = True
  AttackTxt.Visible = False
  AttackTot.Visible = False
  AttackLnd.Visible = False
  AttackWtr.Visible = False
  DefendTxt.Visible = False
  DefendTot.Visible = False
  DefendLnd.Visible = False
  DefendWtr.Visible = False
  Exit Function
End If

CntryName.Caption = MyMap.CountryName(CurrentMouse)
If MyMap.Owner(CurrentMouse) > 0 Then
  CntryName.BackColor = PlayerColorCodes(Player(MyMap.Owner(CurrentMouse)))
  CntryName.ForeColor = PlayerTextColor(Player(MyMap.Owner(CurrentMouse)))
Else
  CntryName.BackColor = Land.BackColor
  CntryName.ForeColor = vbBlack
End If
AttackTot.Caption = Trim(Str(AttackStrength))
DefendTot.Caption = Trim(Str(DefendStrength))
AttackLnd.Caption = "Land: " & Trim(Str(AttLndNum))
AttackWtr.Caption = "Water: " & Trim(Str(AttWtrNum))
DefendLnd.Caption = "Land: " & Trim(Str(DefLndNum))
DefendWtr.Caption = "Water: " & Trim(Str(DefWtrNum))
CntryName.Visible = True
AttackTxt.Visible = True
AttackTot.Visible = True
AttackLnd.Visible = True
AttackWtr.Visible = True
DefendTxt.Visible = True
DefendTot.Visible = True
DefendLnd.Visible = True
DefendWtr.Visible = True

'Now let's tweak what needs to be on and off, etc.

If MyMap.Owner(CurrentMouse) = RightClicker Then
  'The player right-clicked their own country.
  'Don't show the attack strength.
  AttackTxt.Visible = False
  AttackTot.Visible = False
  AttackLnd.Visible = False
  AttackWtr.Visible = False
  'Defend color should be theirs.
  DefendTxt.BackColor = PlayerColorCodes(Player(RightClicker))
  DefendTxt.ForeColor = PlayerTextColor(Player(RightClicker))
ElseIf MyMap.Owner(CurrentMouse) = 0 Then
  'The player right-clicked unclaimed land.
  'Don't show attack or defense, but mention that it's empty.
  AttackTxt.Visible = False
  AttackTot.Visible = False
  AttackLnd.Visible = False
  AttackWtr.Visible = False
  DefendTxt.Visible = False
  DefendTot.Visible = False
  DefendLnd.Caption = "Unclaimed"
  DefendWtr.Visible = False
Else
  'The player right-clicked enemy land.
  'All things are visible, but change the attack and defend text
  'so that the colors are right.
  AttackTxt.BackColor = PlayerColorCodes(Player(RightClicker))
  AttackTxt.ForeColor = PlayerTextColor(Player(RightClicker))
  DefendTxt.BackColor = PlayerColorCodes(Player(MyMap.Owner(CurrentMouse)))
  DefendTxt.ForeColor = PlayerTextColor(Player(MyMap.Owner(CurrentMouse)))
  'If there is no attack strength here, let's get rid of the extra numbers.
  If AttackStrength = 0 Then
    AttackTot.Caption = "----"
    AttackLnd.Visible = False
    AttackWtr.Visible = False
  End If
End If

End Function

Private Sub PlyrNames_Click(Index As Integer)

'Just flash this player's countries!

Dim i As Long

If (PlayerType(Index) > PTYPE_INACTIVE) And (NumOccupied(Index) > 0) Then
  For i = 1 To MyMap.NumberOfCountries
    If MyMap.CountryColor(i) = Player(Index) Then
      'This country belongs to the player whose name we clicked.  Flash it.
      LongFlash (i)
    End If
  Next i
  DrawMap
End If

End Sub

Private Sub CntryName_Click()

Dim i As Long
Dim j As Integer
Dim Cntry As Long
Dim FoundIt As Boolean

'Compensate for two right-clicks in a row...
LastRightClick = Abs(LastRightClick)

'Allow the player to change the name of this country.
'Leave if not our country.
If (PlayerType(Turn) <> PTYPE_HUMAN) Then
  Land.Oopsie.Caption = "You can only rename your countries during your turn."
  Call SNDBooBoo
  Exit Sub
End If

'Leave if not a country or water at all.
If (LastRightClick < 1) Then
  Land.Oopsie.Caption = "Only land and water can be renamed."
  Call SNDBooBoo
  Exit Sub
End If

'If this is water, all of its coastline must belong to the current player.
FoundIt = False
If (LastRightClick > TILEVAL_COASTLINE) Then
  'See if the current player owns all coastline near this water mass.
  For i = 1 To MyMap.NumberOfCountries
    If FoundIt = True Then Exit For
    For j = 1 To MyMap.MaxNeighbors
      Cntry = MyMap.Neighbors(i, j)
      If (Cntry = 0) Then Exit For  'Done with this country.
      If (Cntry = LastRightClick) Then   'It's water.
        If (MyMap.Owner(i) <> Turn) Then
          'This is a country we don't own which borders the right-clicked one.
          FoundIt = True
          Exit For
        End If
      End If
    Next j
  Next i
  'Now see if we found a country that borders us that we don't own yet.
  If FoundIt = True Then
    Land.Oopsie.Caption = "To rename a body of water, you must own all of its coastline."
    Call SNDBooBoo
    Exit Sub
  Else
    'We can rename this one!  Populate the rename dialog....
    RenameCountry.OldName.Caption = MyMap.WaterName(LastRightClick - 1000)
    RenameCountry.OldName.BackColor = RenameCountry.BackColor
    RenameCountry.NewName = ""
    '....and show it.
    RenameCountry.Show vbModal
    Exit Sub
  End If
End If

'At this point, we are trying to rename a country.  Leave if not our country.
If (MyMap.Owner(LastRightClick) <> Turn) Then
  Land.Oopsie.Caption = "You can only rename your own countries."
  Call SNDBooBoo
  Exit Sub
Else
  'Populate the rename dialog....
  RenameCountry.OldName.Caption = MyMap.CountryName(LastRightClick)
  RenameCountry.OldName.BackColor = PlayerColorCodes(MyMap.CountryColor(LastRightClick))
  RenameCountry.NewName = ""
  '....and show it.
  RenameCountry.Show vbModal
End If


End Sub

Private Sub ResignBut_Click()

Dim Resp As Long
Dim a As Long
Dim i As Integer

'Leave if this isn't a human's first turn phase.
If (Phase <> 1) Or (PlayerType(Turn) <> PTYPE_HUMAN) Then Exit Sub

'See if the *last* player is trying to resign.  There's a window where
'this could happen if you're just clickin' stuff.
a = 0
For i = 1 To 6
  If NumOccupied(i) > 0 Then a = a + 1
Next i
If a = 1 Then Exit Sub  'We're the only one left!  Wait until game over hits.

'The player has resigned!  Let's check to make sure they meant it...
Resp = MsgBox(PlayerName(Turn) & ": Are you sure you want to resign?", vbYesNo, "Resign")
'TODO: better text if network player.

If Resp = vbNo Then Exit Sub   'I didn't think so.

'Perform the resign.
Call ResignProc(Turn)
'Send this signal to others if necessary.
Call SendResignToNetwork(Turn, Phase, 0)

ResignBut.Visible = False

'The GameTimer will detect that the resign has occured and pass the turn.

End Sub

Public Sub ResignProc(Turn As Integer)

Dim a As Long

'Now we need to resign the player.  Their countries either stay put
'or go away based on the config settings.
For a = 1 To MyMap.NumberOfCountries
  If MyMap.Owner(a) = Turn Then
    'We need to do something with this country.
    Select Case CFGConquer
      Case 2, 4:
        'All this player's countries turn neutral again.
        'Kill ports, leave troops.
        MyMap.Owner(a) = 0
        MyMap.CountryType(a) = 0
        'Assign it an 'unoccupied' color.
        MyMap.CountryColor(a) = CFGUnoccupiedColor
        'If 'Enemy is completely eradicated', then remove troops.
        If CFGConquer = 4 Then
          MyMap.TroopCount(a) = 0
        End If
      Case 1, 3:
        'Countries stay owned by the defunct player.
        'Basically, we do nothing here!  The remaining players
        'will need to attack and conquer the countries to claim them.
        'Since there is no 'victor' here, we'll leave them alone.
        'We do need to kill the HQ, though...
        If (MyMap.CountryType(a) And 2) = 2 Then
          MyMap.CountryType(a) = MyMap.CountryType(a) And 253
        End If
      Case 5:
        'Chaos erupts!
        If MyNetworkRole <> NW_CLIENT Then
          'First, kill all ports and HQs.
          MyMap.CountryType(a) = 0
          'There is a 50% chance of going neutral, and a
          '50% chance of staying loyal.
          If Rnd(1) < 0.5 Then
            'Country goes neutral.
            MyMap.Owner(a) = 0
            MyMap.CountryColor(a) = CFGUnoccupiedColor
          'Else it stays loyal, do nothing.
          End If
          'Now see if we kill troops in it.  40% chance.
          If Rnd(1) < 0.4 Then
            'Kill 'em all!
            MyMap.TroopCount(a) = 0
          End If
        End If
    End Select
  End If
Next a

NumOccupied(Turn) = -1  'This effectively disables the player.

'Update our statistics...
StatScreen.UpdateResignedStats (Turn)

'Now send all map data to the clients if we are the server.
If (MyNetworkRole = NW_SERVER) And (CFGConquer = 5) Then SendNewCountryData

'Refresh, since lots of countries may have changed...
Call Land.RedrawScreen

End Sub

Private Sub TroopMoveDn_Click()

'Decrement the troop entry prompt by one.
If Val(TroopMoveNum.Text) > 1 Then
  TroopMoveNum.Text = Trim(Str(Val(TroopMoveNum.Text) - 1))
End If

End Sub

Private Sub TroopMoveNum_Change()

'This sub is called whenever the number in the troop entry prompt changes.
'We will verify that it is between 1 and the max troops for the selected
'country and adjust it if it isn't.  Basically, the number in the box
'will always be valid.  No error checking is necessary elsewhere!

If TroopMoveNum.Visible = False Then Exit Sub

'Check for non-numeric characters.
Dim q As Integer
For q = 1 To Len(TroopMoveNum.Text)
  If Asc(Mid(TroopMoveNum.Text, q, 1)) < 48 Or Asc(Mid(TroopMoveNum.Text, q, 1)) > 57 Then
    TroopMoveNum.Text = "1"
    Exit Sub
  End If
Next q

'Check for too high.
If (Val(TroopMoveNum.Text) > MyMap.TroopCount(TroopMoveSrc)) Then
  TroopMoveNum.Text = Trim(Str(MyMap.TroopCount(TroopMoveSrc)))
  Exit Sub
End If

'Check for too low.
If TroopMoveNum.Text = "" Or Val(TroopMoveNum.Text) = 0 Then
  TroopMoveNum.Text = "1"
  Exit Sub
End If

End Sub

Private Sub TroopMoveQtr_Click()

'Put 25% of troops in the troop entry prompt.
Dim TempNum As Long

TempNum = Int(MyMap.TroopCount(TroopMoveSrc) * 0.25)
If TempNum = 0 Then TempNum = 1

TroopMoveNum.Text = Trim(Str(TempNum))

End Sub

Private Sub TroopMoveHlf_Click()

'Put 50% of troops in the troop entry prompt.
Dim TempNum As Long

TempNum = Int(MyMap.TroopCount(TroopMoveSrc) * 0.5)
If TempNum = 0 Then TempNum = 1

TroopMoveNum.Text = Trim(Str(TempNum))

End Sub

Private Sub TroopMove3Qt_Click()

'Put 75% of troops in the troop entry prompt.
Dim TempNum As Long

TempNum = Int(MyMap.TroopCount(TroopMoveSrc) * 0.75)
If TempNum = 0 Then TempNum = 1

TroopMoveNum.Text = Trim(Str(TempNum))

End Sub

Private Sub TroopMoveAll_Click()

'Put 100% of troops in the troop entry prompt.
TroopMoveNum.Text = Trim(Str(MyMap.TroopCount(TroopMoveSrc)))

End Sub

Private Sub TroopMoveUp_Click()

'Increment the troop entry prompt by one.
If Val(TroopMoveNum.Text) < MyMap.TroopCount(TroopMoveSrc) Then
  TroopMoveNum.Text = Trim(Str(Val(TroopMoveNum.Text) + 1))
End If

End Sub

Private Sub CancelBut_Click()

'Pressing this button will allow the player to choose a new source country
'during the troop movement phase.
If Phase = 4 Then Phase = 3

'Also clear our double-click prevention flags.
AlreadyMovedTroopsThisPhase = False
AlreadyChoseTroopsThisPhase = False

'And get rid of the troop movement controls.
SetUpPlayerControls

End Sub

Private Sub MenuAniSpeed_Click()

Dim TempMode As Integer

TempMode = GameMode

If GameMode = GM_TITLE_SCREEN Then
  GameMode = GM_TITLE_DIALOG_OPEN
Else
  GameMode = GM_DIALOG_OPEN
End If

GfxOptions.Show vbModeless
GameMode = TempMode  'Restore the old mode.

'Draw the map.
DrawMap
End Sub

Private Sub InitPlayerData()

'Initialize the player settings.
Player(1) = 1
Player(2) = 3
Player(3) = 5
Player(4) = 9
Player(5) = 7
Player(6) = 4
PlayerType(1) = PTYPE_HUMAN
PlayerType(2) = PTYPE_COMPUTER
PlayerType(3) = PTYPE_COMPUTER
PlayerType(4) = PTYPE_COMPUTER
PlayerType(5) = PTYPE_COMPUTER
PlayerType(6) = PTYPE_COMPUTER
PlayerName(1) = "Player 1"
PlayerName(2) = "Player 2"
PlayerName(3) = "Player 3"
PlayerName(4) = "Player 4"
PlayerName(5) = "Player 5"
PlayerName(6) = "Player 6"

PlayerColors(1) = "Purple"
PlayerColors(2) = "Blue"
PlayerColors(3) = "Yellow"
PlayerColors(4) = "Orange"
PlayerColors(5) = "Red"
PlayerColors(6) = "Gray"
PlayerColors(7) = "Light Blue"
PlayerColors(8) = "Pink"
PlayerColors(9) = "Green"
PlayerColors(10) = "Light Green"
PlayerColors(11) = "Brown"
PlayerColors(12) = "White"

PlayerColorCodes(1) = RGB(204, 51, 255)
PlayerColorCodes(2) = RGB(0, 153, 204)
PlayerColorCodes(3) = RGB(255, 255, 51)
PlayerColorCodes(4) = RGB(255, 153, 0)
PlayerColorCodes(5) = RGB(204, 51, 0)
PlayerColorCodes(6) = RGB(102, 102, 102)
PlayerColorCodes(7) = RGB(0, 255, 255)
PlayerColorCodes(8) = RGB(255, 102, 153)
PlayerColorCodes(9) = RGB(0, 153, 0)
PlayerColorCodes(10) = RGB(102, 255, 51)
PlayerColorCodes(11) = RGB(153, 102, 0)
PlayerColorCodes(12) = RGB(255, 255, 255)

PlayerTextColor(1) = vbBlack
PlayerTextColor(2) = vbBlack
PlayerTextColor(3) = vbBlack
PlayerTextColor(4) = vbBlack
PlayerTextColor(5) = vbWhite
PlayerTextColor(6) = vbWhite
PlayerTextColor(7) = vbBlack
PlayerTextColor(8) = vbBlack
PlayerTextColor(9) = vbWhite
PlayerTextColor(10) = vbBlack
PlayerTextColor(11) = vbWhite
PlayerTextColor(12) = vbBlack

'Also initialize anything else that shouldn't be a 0 or FALSE on start.
CFGExplosions = 1
CFGWaves = 1
CFGUnoccupiedColor = 6
CFGFlashing = 1
CFGPrompt = 1

End Sub

Private Function PickedHQYet() As Boolean

Dim i As Integer

'Return False if all players haven't picked HQ yet during auto HQ selection.
PickedHQYet = True
If CFGHQSelect = 2 Then
  For i = 1 To 6
    If (NumOccupied(i) = 0) And PlayerType(i) > PTYPE_INACTIVE Then
      PickedHQYet = False
      Exit Function
    End If
  Next i
End If

End Function

Private Sub SetupForm()

Dim WinWidth As Long
Dim WinHeight As Long
Dim LetsResize As Boolean

'This function just sets up our form's parameters, dimensions, and the like.
'We return TRUE if we had to change the dimensions, and FALSE if
'we were already at this resolution.

'First, set up our form's dimensions based on the menu setting.
Select Case CFGResolution
  Case 1:    '640x480
    WinWidth = 640
    WinHeight = 450
    MyMap.SetDimensions XDIM640x480, YDIM640x480
  Case 2:    '800x600
    WinWidth = 800
    WinHeight = 570
    MyMap.SetDimensions XDIM800x600, YDIM800x600
  Case 3:    '1024x768
    WinWidth = 1024
    WinHeight = 738
    MyMap.SetDimensions XDIM1024x768, YDIM1024x768
End Select

'See if we need to resize this form later.
LetsResize = False
If (WinWidth * Screen.TwipsPerPixelX) <> Land.Width Then
  'We'll be resizing and centering this form.
  LetsResize = True
End If

Wide = (MyMap.Xsize * 8)
Tall = (MyMap.Ysize * 8)

Ocean.Width = Wide * Screen.TwipsPerPixelX
Ocean.Height = Tall * Screen.TwipsPerPixelY
MapBuffer.Width = Wide * Screen.TwipsPerPixelX
MapBuffer.Height = Tall * Screen.TwipsPerPixelY
PicBuffer.Width = Wide * Screen.TwipsPerPixelX
PicBuffer.Height = Tall * Screen.TwipsPerPixelY

'Set up window position and dimensions.
Land.Width = WinWidth * Screen.TwipsPerPixelX
Land.Height = WinHeight * Screen.TwipsPerPixelY
Land.ScaleWidth = Land.Width
Land.ScaleHeight = Land.Height
Land.ScaleMode = 1

'Add other controls.
PassTurnBut.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
PassTurnBut.Top = 1 * Screen.TwipsPerPixelY
PassTurnBut.Visible = False

PlyrFrame.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
PlyrFrame.Top = 22 * Screen.TwipsPerPixelY
PlyrFrame.Visible = False

CntryFrame.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
CntryFrame.Top = 225 * Screen.TwipsPerPixelY
CntryFrame.Visible = False

TroopMoveLbl.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
TroopMoveLbl.Top = (WinHeight - 157) * Screen.TwipsPerPixelY
TroopMoveLbl.Visible = False

TroopMoveNum.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
TroopMoveNum.Top = (WinHeight - 140) * Screen.TwipsPerPixelY
TroopMoveNum.Text = "1"
TroopMoveNum.Visible = False

TroopMoveUp.Left = (WinWidth - 32) * Screen.TwipsPerPixelX
TroopMoveUp.Top = (WinHeight - 148) * Screen.TwipsPerPixelY
TroopMoveUp.Visible = False

TroopMoveDn.Left = (WinWidth - 32) * Screen.TwipsPerPixelX
TroopMoveDn.Top = (WinHeight - 131) * Screen.TwipsPerPixelY
TroopMoveDn.Visible = False

TroopMoveQtr.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
TroopMoveQtr.Top = (WinHeight - 112) * Screen.TwipsPerPixelY
TroopMoveQtr.Visible = False

TroopMoveHlf.Left = (WinWidth - 69) * Screen.TwipsPerPixelX
TroopMoveHlf.Top = (WinHeight - 112) * Screen.TwipsPerPixelY
TroopMoveHlf.Visible = False

TroopMove3Qt.Left = (WinWidth - 51) * Screen.TwipsPerPixelX
TroopMove3Qt.Top = (WinHeight - 112) * Screen.TwipsPerPixelY
TroopMove3Qt.Visible = False

TroopMoveAll.Left = (WinWidth - 32) * Screen.TwipsPerPixelX
TroopMoveAll.Top = (WinHeight - 112) * Screen.TwipsPerPixelY
TroopMoveAll.Visible = False

CancelBut.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
CancelBut.Top = (WinHeight - 87) * Screen.TwipsPerPixelY
CancelBut.Visible = False

ResignBut.Left = (WinWidth - 88) * Screen.TwipsPerPixelX
ResignBut.Top = 0 * Screen.TwipsPerPixelY
ResignBut.Visible = False

Commentary.Left = XOFFSET * Screen.TwipsPerPixelX
Commentary.Top = Land.ScaleHeight - (25 * Screen.TwipsPerPixelY)
Commentary.Caption = "Choose game parameters from the Options menu, and landmass parameters from the Terraform menu."
Commentary.Width = (Wide - 4) * Screen.TwipsPerPixelX
Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack

Oopsie.Left = XOFFSET * Screen.TwipsPerPixelX
Oopsie.Top = Land.ScaleHeight - (42 * Screen.TwipsPerPixelY)
Oopsie.Caption = "Welcome to Fracas."
Oopsie.Width = (Wide - 4) * Screen.TwipsPerPixelX
Oopsie.BackColor = Land.BackColor

If LetsResize Then CenterOurForm

End Sub

Private Sub CenterOurForm()

Land.Left = (Screen.Width / 2) - (Land.Width / 2)

If (CFGResolution = MaxAllowedResolution) And ((Screen.Height \ Screen.TwipsPerPixelY) < 800) Then
  Land.Top = 0
Else
  Land.Top = (Screen.Height / 2) - (Land.Height / 2)
End If

End Sub

Public Sub ChangeRes()

'Things to do when the map resolution changes.
Set MyMap = New Map

'DoEvents
  
'If the main form is minimized, get it back up!
If Land.WindowState = vbMinimized Then
  Land.WindowState = vbNormal
  Land.SetFocus
End If

SetupForm
'DoEvents
MyMap.SetupTitleScreen
Call DrawBkg
'DoEvents
Call DrawMap
Land.Refresh
  
End Sub

Private Sub DetermineMaxRes()

'Only enable resolutions less than or equal to what we have currently.
If (Screen.Width / Screen.TwipsPerPixelX) < 800 Then
  MaxAllowedResolution = 1
ElseIf (Screen.Width / Screen.TwipsPerPixelX) < 1024 Then
  MaxAllowedResolution = 2
Else
  MaxAllowedResolution = 3
End If

End Sub

Private Sub FixResolution()

Dim i As Integer

'This sub verifies that our resolution is good.
If CFGResolution > MaxAllowedResolution Then
  CFGResolution = MaxAllowedResolution
  For i = 1 To 3
    MenuResolution(i).Checked = False
  Next i
  MenuResolution(CFGResolution).Checked = True
End If

End Sub

Public Sub SetupMenusForGameOver()

Dim i As Integer

'Get rid of the troop movement entry tool and other player controls.
Commentary.BackColor = Land.BackColor
Commentary.ForeColor = vbBlack
PassTurnBut.Visible = False
TroopMoveNum.Visible = False
TroopMoveUp.Visible = False
TroopMoveDn.Visible = False
TroopMoveLbl.Visible = False
TroopMoveQtr.Visible = False
TroopMoveHlf.Visible = False
TroopMove3Qt.Visible = False
TroopMoveAll.Visible = False
CancelBut.Visible = False
ResignBut.Visible = False
PlyrFrame.Visible = False
CntryFrame.Visible = False

'Set up menu items for when a game is not being played.
MenuNew.Enabled = True
MenuLoad.Enabled = True
MenuGetOptionsFromMapFile.Enabled = True
MenuSame.Enabled = False
MenuHiScores.Enabled = False
If LastMapPath <> "" Then
  MenuSame.Enabled = True
  MenuHiScores.Enabled = True
End If
MenuLoadSavedGame.Enabled = True
MenuJoinGame.Enabled = True
MenuChat.Enabled = False
MenuSave.Enabled = False
MenuSaveGame.Enabled = False
MenuPlayers.Enabled = True
MenuAbortGame.Enabled = False
MenuStats.Enabled = False
For i = 0 To MAX_MENU_ITEMS
  If i <= 7 And i > 0 Then
    Menu1st(i).Enabled = True
  End If
  If i <= 6 And i > 0 Then
    MenuSize(i).Enabled = True
  End If
  If i <= 5 Then
    MenuBonus(i).Enabled = True
  End If
  If i <= 5 And i > 0 Then
    MenuInitTroops(i).Enabled = True
    MenuPct(i).Enabled = True
    MenuConquer(i).Enabled = True
  End If
  If i <= 4 And i > 0 Then
    MenuPorts(i).Enabled = True
    MenuShips(i).Enabled = True
    MenuLakeSize(i).Enabled = True
  End If
  If i <= 3 And i > 0 Then
    MenuRandom(i).Enabled = True
    MenuInitTroopCts(i).Enabled = True
    MenuIslands(i).Enabled = True
    MenuShape(i).Enabled = True
    MenuProp(i).Enabled = True
    MenuResolution(i).Enabled = False  'Turn on the ones we want below.
  End If
  If i <= 2 And i > 0 Then
    MenuHQSelect(i).Enabled = True
  End If
Next i

'Disable no starting troops if we have no bonus troops.
If MenuBonus(0).Checked = True Then
  MenuInitTroops(1).Enabled = False
Else
  MenuInitTroops(1).Enabled = True
End If

'Disable no bonus troops if we start with none.
If MenuInitTroops(1).Checked = True Then
  MenuBonus(0).Enabled = False
Else
  MenuBonus(0).Enabled = True
End If

'Only enable resolutions less than or equal to what we have currently.
For i = 1 To MaxAllowedResolution
  MenuResolution(i).Enabled = True
Next i

End Sub

Private Sub SetupMenusForInGame()

Dim i As Integer

'Set up menu items for in-game use.
MenuNew.Enabled = False
MenuSame.Enabled = False
MenuLoad.Enabled = False
MenuGetOptionsFromMapFile.Enabled = False
MenuLoadSavedGame = False
MenuSave.Enabled = True
MenuJoinGame.Enabled = False
MenuPlayers.Enabled = False
MenuAbortGame.Enabled = True
MenuHiScores.Enabled = True
MenuChat.Enabled = False
If MyNetworkRole <> NW_NONE Then
  MenuHiScores.Enabled = False
  MenuChat.Enabled = True
End If
MenuStats.Enabled = True
For i = 0 To MAX_MENU_ITEMS
  If i <= 7 And i > 0 Then
    Menu1st(i).Enabled = False
  End If
  If i <= 5 Then
    MenuBonus(i).Enabled = False
  End If
  If i <= 6 And i > 0 Then
    MenuSize(i).Enabled = False
  End If
  If i <= 5 And i > 0 Then
    MenuInitTroops(i).Enabled = False
    MenuPct(i).Enabled = False
    MenuConquer(i).Enabled = False
  End If
  If i <= 4 And i > 0 Then
    MenuPorts(i).Enabled = False
    MenuShips(i).Enabled = False
    MenuLakeSize(i).Enabled = False
  End If
  If i <= 3 And i > 0 Then
    MenuRandom(i).Enabled = False
    MenuInitTroopCts(i).Enabled = False
    MenuIslands(i).Enabled = False
    MenuResolution(i).Enabled = False
    MenuShape(i).Enabled = False
    MenuProp(i).Enabled = False
  End If
  If i <= 2 And i > 0 Then
    MenuHQSelect(i).Enabled = False
  End If
Next i

End Sub

Public Sub SetUpPlayerControls()

Select Case PlayerType(Turn)
Case PTYPE_HUMAN:
  'This is a human's turn.
  'See if they have any countries at all.
  If (NumOccupied(Turn) = 0) And (CFGHQSelect = 1) Then
    'Nope, then this is their FIRST turn.
    PassTurnBut.Visible = False
    CntryFrame.Visible = True
    Exit Sub
  End If
  
  Select Case Phase
  Case 1:
    'Reinforcement phase.
    PassTurnBut.Visible = False
    TroopMoveNum.Visible = False
    TroopMoveUp.Visible = False
    TroopMoveDn.Visible = False
    TroopMoveLbl.Visible = False
    TroopMoveQtr.Visible = False
    TroopMoveHlf.Visible = False
    TroopMove3Qt.Visible = False
    TroopMoveAll.Visible = False
    CancelBut.Visible = False
    CntryFrame.Visible = True
    Exit Sub
  Case 2:
    'Action phase.
    PassTurnBut.Visible = True
    TroopMoveNum.Visible = False
    TroopMoveUp.Visible = False
    TroopMoveDn.Visible = False
    TroopMoveLbl.Visible = False
    TroopMoveQtr.Visible = False
    TroopMoveHlf.Visible = False
    TroopMove3Qt.Visible = False
    TroopMoveAll.Visible = False
    CancelBut.Visible = False
    ResignBut.Visible = False
    CntryFrame.Visible = True
    Exit Sub
  Case 3:
    'Troop movement source select phase.
    PassTurnBut.Visible = True
    TroopMoveNum.Visible = False
    TroopMoveUp.Visible = False
    TroopMoveDn.Visible = False
    TroopMoveLbl.Visible = False
    TroopMoveQtr.Visible = False
    TroopMoveHlf.Visible = False
    TroopMove3Qt.Visible = False
    TroopMoveAll.Visible = False
    CancelBut.Visible = False
    ResignBut.Visible = False
    CntryFrame.Visible = True
    Exit Sub
  Case 4:
    'Troop movement destination select phase.
    PassTurnBut.Visible = True
    TroopMoveNum.Visible = True
      'Set the focus there and highlight the text so we can type immediately.
      TroopMoveNum.SetFocus
      TroopMoveNum.SelStart = 0
      TroopMoveNum.SelLength = 10
    TroopMoveUp.Visible = True
    TroopMoveDn.Visible = True
    TroopMoveLbl.Visible = True
    TroopMoveQtr.Visible = True
    TroopMoveHlf.Visible = True
    TroopMove3Qt.Visible = True
    TroopMoveAll.Visible = True
    CancelBut.Visible = True
    ResignBut.Visible = False
    If CFGResolution = 1 Then
      CntryFrame.Visible = False
    Else
      CntryFrame.Visible = True
    End If
    Exit Sub
  End Select
  
Case PTYPE_COMPUTER, PTYPE_NETWORK, PTYPE_SERVER:
  'This is a computer's turn!
  'Get rid of all the onscreen controls...
  PassTurnBut.Visible = False
  TroopMoveNum.Visible = False
  TroopMoveUp.Visible = False
  TroopMoveDn.Visible = False
  TroopMoveLbl.Visible = False
  TroopMoveQtr.Visible = False
  TroopMoveHlf.Visible = False
  TroopMove3Qt.Visible = False
  TroopMoveAll.Visible = False
  CancelBut.Visible = False
  ResignBut.Visible = False

End Select

End Sub

Public Sub CalculateTotals()

'This sub recalculates the values that should be in the box to the right.
'Used when setting up a client so we don't have to send these over the wire.
'Also used after resending all country data after a random event.

Dim i As Long
Dim j As Integer

For j = 1 To MAX_PLAYERS
  If NumOccupied(j) <> -1 Then
    NumOccupied(j) = 0
    NumTroops(j) = 0
  End If
Next j

For i = 1 To MyMap.NumberOfCountries
  j = MyMap.Owner(i)
  If j > 0 Then
    If NumOccupied(j) <> -1 Then
      NumOccupied(j) = NumOccupied(j) + 1
      NumTroops(j) = NumTroops(j) + MyMap.TroopCount(i)
    End If
  End If
Next i

End Sub

Public Sub PaintGame()
    'wrapper for the Paint event of the form to force a redraw
    Call Form_Paint
End Sub
