Attribute VB_Name = "Declarations"
Option Explicit
Option Base 1

Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Const SRCCOPY        As Long = &HCC0020          ' (DWORD) dest = source
Public Const SRCINVERT      As Long = &H660046          ' (DWORD) dest = source XOR dest
Public Const SRCPAINT       As Long = &HEE0086          ' (DWORD) dest = source OR dest
Public Const SRCAND         As Long = &H8800C6          ' (DWORD) dest = source AND dest

Public Const VersionStr = "Fracas 2.0 BETA 1.5 by Jason Merlo"
Public Const WebStr = "http://www.smozzie.com"
Public Const EmailStr = "jmerlo@austin.rr.com"

Public GameMode As Integer
Public Const GM_TITLE_SCREEN        As Integer = 1
Public Const GM_BUILDING_MAP        As Integer = 2
Public Const GM_BUILDING_TITLE      As Integer = 3
Public Const GM_DIALOG_OPEN         As Integer = 4
Public Const GM_TITLE_DIALOG_OPEN   As Integer = 5
Public Const GM_GAME_ACTIVE         As Integer = 6

'RollOver tells us that a map exists on the screen and
'we can therefore interact with it.
'Public RollOver As Boolean
'Public GameOver As Boolean
'Public TitleShown As Boolean
'Public MapBeingBuilt As Boolean

'MapMade tells us that at least one map has been made during this session.
Public MapMade As Boolean

'These parameters are used to parse menu settings.
Public MinLakeSize As Integer
Public MaxCountrySize As Integer
Public CFGCountrySize As Integer
Public CFGLandPct As Integer
Public CFGIslands As Integer
Public CFGLakeSize As Integer
Public CFGProportion As Integer
Public CFGShape As Integer
Public CFGBorders As Integer
Public CFGInitTroopPl As Integer
Public CFGInitTroopCt As Integer
Public CFGBonusTroops As Integer
Public CFGShips As Integer
Public CFGPorts As Integer
Public CFGConquer As Integer
Public CFGAISpeed As Integer
Public CFGEvents As Integer
Public CFG1st As Integer
Public CFGHQSelect As Integer
Public CFGSound As Integer
Public CFGExplosions As Integer          'Toggle for ball engine.
Public CFGWaves As Integer               'Toggle for water animations.
Public CFGUnoccupiedColor As Integer     'Color code of empty countries.
Public CFGFlashing As Integer            'Do we flash countries while moving?
Public CFGResolution As Integer          'Horizontal resolution we're currently at.
Public CFGPrompt As Integer              'Prompt after the game is over?
Public MaxAllowedResolution As Integer   'Which resolution is the max.
Public u As Integer
Public LandPct As Double
Public PropPct As Double
Public ShapePct As Double
Public CoastPctKeep As Double
Public IslePctKeep As Double
Public ShipPct As Single
Public Done As Boolean

'These are the screen dimensions.
Public Wide As Long
Public Tall As Long

Public Const XDIM640x480    As Integer = 66
Public Const YDIM640x480    As Integer = 45 ' Add 3 for 1.5 compatibility.
Public Const XDIM800x600    As Integer = 86
Public Const YDIM800x600    As Integer = 60 ' Add 3 for 1.5 compatibility.
Public Const XDIM1024x768   As Integer = 114
Public Const YDIM1024x768   As Integer = 80

'Make some map variables public.
Public MyMap As Map
Public LastRightClick As Long

'Number of times to loop through menu routines.
Public Const MAX_MENU_ITEMS As Integer = 7

'Gameplay constants.
Public Const MAX_COUNTRY_CAPACITY As Integer = 999

'Constants used to get around in the map builder.
Public Const TILEVAL_COASTLINE          As Integer = 999
Public Const MIN_COUNTRY_SIZE           As Integer = 30  'Max number of blocks on the map per country.
Public Const MAX_COUNTRY_SIZE           As Integer = 150
Public Const MAX_PROP_PCT               As Single = 0.7
Public Const MAX_COUNTRY_TRIES          As Integer = 600
Public Const MAX_TRIES_TO_PLACE_COUNTRY As Integer = 100
Public Const UNUSABLE_GRID              As Integer = -37
Public Const MAXIMUM_NEIGHBORS          As Integer = 20
Public Const UP                         As Integer = 1
Public Const DOWN                       As Integer = 2
Public Const LEFTY                      As Integer = 3
Public Const RIGHTY                     As Integer = 4

Public Const GFX_NUM_LAND_PIECES        As Integer = 19
Public Const GFX_NUM_LAND_PIECE_SETS    As Integer = 4
Public Const GFX_NUM_PIER_PIECES        As Integer = 14
Public Const GFX_NUM_PIER_PIECE_SETS    As Integer = 3
Public Const GFX_GRID_HALF              As Integer = 4
Public Const GFX_GRID                   As Integer = 8
Public Const GFX_GRID_X2                As Integer = 16

Public Const GFX_PIER_X_BASE            As Integer = 29
Public Const GFX_DIGIT_X_BASE           As Integer = 20
Public Const GFX_EDIT_PIECE_X_BASE      As Integer = 15
Public Const GFX_BLACK_X_COORD          As Long = (4 * GFX_GRID)
Public Const GFX_WHITE_X_COORD          As Long = (17 * GFX_GRID)
Public Const GFX_WORK_X_COORD           As Long = (18 * GFX_GRID)
Public Const GFX_HQ_X_COORD             As Long = (19 * GFX_GRID)

Public Const GFX_MASK_Y_LINE            As Integer = 12
Public Const GFX_SUPP_Y_LINE            As Integer = 14
Public Const GFX_SUPP_MASK_Y_LINE       As Integer = 15

Public Const GFX_ICONS_X                As Integer = 5
Public Const GFX_ICONS_Y                As Integer = 13

Public Const GFX_LAND_MASK_Y_COORD      As Long = (GFX_MASK_Y_LINE * GFX_GRID)
Public Const GFX_HQ_Y_COORD             As Long = (GFX_SUPP_Y_LINE * GFX_GRID)
Public Const GFX_WHITE_Y_COORD          As Long = (GFX_SUPP_MASK_Y_LINE * GFX_GRID)

'Some directions, used for borders...
Public Const DIR_UP                     As Integer = 0
Public Const DIR_RIGHT                  As Integer = 1

'These are the offsets (in pixels) for the map within the form container.
Public Const XOFFSET                    As Integer = 10
Public Const YOFFSET                    As Integer = 1

'The maximum number of balls that we can track.
Public Const BALLMAX                    As Integer = 500

'The maximum number of waves in the ocean at one time.
Public Const WAVEMAX                    As Integer = 100
Public Const OCEAN_X_SIZE               As Integer = 200
Public Const OCEAN_Y_SIZE               As Integer = 150

'The gravitational constant (in pixels/timer cycle^2)
Public Const GRAVITY                    As Integer = 1

'These constants contain properties for each type of explosion used.
'Note that some of these are multipliers, not absolutes.
'Little splashes in the water.
Public Const SPLASH_COUNT               As Integer = 18
Public Const SPLASH_INTENSITY           As Integer = 15
Public Const SPLASH_SPREAD              As Integer = 20
Public Const SPLASH_ELASTIC             As Integer = 0
Public Const SPLASH_SIZE                As Integer = 2
Public Const SPLASH_COLOR               As Integer = 0
'A small explosion.
Public Const SMALL_COUNT                As Single = 0.2
Public Const SMALL_INTENSITY            As Single = 0.03
Public Const SMALL_SPREAD               As Integer = 25
Public Const SMALL_ELASTIC              As Integer = 50
Public Const SMALL_SIZE                 As Integer = 2
'A medium explosion.
Public Const MED_COUNT                  As Single = 0.3
Public Const MED_INTENSITY              As Single = 0.05
Public Const MED_SPREAD                 As Integer = 30
Public Const MED_ELASTIC                As Integer = 55
Public Const MED_SIZE                   As Integer = 3
'A big explosion.
Public Const BIG_COUNT                  As Single = 0.4
Public Const BIG_INTENSITY              As Single = 0.07
Public Const BIG_SPREAD                 As Integer = 35
Public Const BIG_ELASTIC                As Integer = 60
Public Const BIG_SIZE                   As Integer = 4
'Bonus twinkles for one country.
Public Const BONUS_COUNT                As Integer = 120
Public Const BONUS_INTENSITY            As Integer = 80
Public Const BONUS_SPREAD               As Integer = 25
Public Const BONUS_ELASTIC              As Integer = 0
Public Const BONUS_SIZE                 As Integer = 7
'Small bonus twinkles for many countries at once.
Public Const SMBONUS_COUNT              As Integer = 20
Public Const SMBONUS_INTENSITY          As Integer = 85
Public Const SMBONUS_SPREAD             As Integer = 19
Public Const SMBONUS_ELASTIC            As Integer = 0
Public Const SMBONUS_SIZE               As Integer = 9

'The PlayerColors array contains text versions of each player color
'for display purposes.  The PlayerColorCodes array contains the color itself.
Public PlayerColors(12) As String
Public PlayerColorCodes(12) As Long
Public PlayerTextColor(12) As Long

'The Player array contains the color code of each player,
'used for identification purposes.  The PlayerName is each players name.
Public Const MAX_PLAYERS As Integer = 6
Public Player(MAX_PLAYERS) As Integer
Public PlayerName(MAX_PLAYERS) As String

'The PlayerType array tells whether or not this player is a computer.
'3 = Networked Human, 2 = Computer Controlled, 1 = Local Human, 0 = Inactive.
Public PlayerType(MAX_PLAYERS) As Integer
Public TempPlayerType(MAX_PLAYERS) As Integer   'Used by clients for saving and displaying icons.
Public Const PTYPE_INACTIVE As Integer = 0
Public Const PTYPE_HUMAN As Integer = 1
Public Const PTYPE_COMPUTER As Integer = 2
Public Const PTYPE_NETWORK As Integer = 3
Public Const PTYPE_NET_AVAIL As Integer = 4 'Only used when clients are choosing color.
Public Const PTYPE_SERVER As Integer = 5 'Used by clients.  The server controls this player.

'The Personality array tells which computer personality this computer
'player is.
Public Personality(MAX_PLAYERS) As Integer

'Turn contains the player whose turn it is (1-6).
'Phase contains the phase of the current turn:
'1 = receive/place troops
'2 = attack
'3,4 = move troops
Public TurnCounter As Long

'These are used to ensure that troop additions, actions, and passes only happen once.
Public AlreadyAddedTroopsThisPhase As Boolean
Public AlreadyPerformedActionThisPhase As Boolean
Public AlreadyChoseTroopsThisPhase As Boolean
Public AlreadyMovedTroopsThisPhase As Boolean
Public AlreadyPassedThisPhase As Boolean
Public AlreadyCheckedForMaxedOutCountries As Boolean

'WonTurn is used to capture the winner of the game.
'Since the turns and phases are dynamic, we need to be able to
'freeze the number.
Public WonTurn As Integer
Public WonGame As Boolean

'NumOccupied contains the number of countries owned by each player.
'NumTroops contains the total number of troops owned by each player.
'Used for AI and for tiebreaking situations.
Public NumOccupied(MAX_PLAYERS) As Long
Public NumTroops(MAX_PLAYERS) As Long

'NumPlayers contains the number of people playing this game (2-6).
'NumNetworkPlayers is the number of networked humans the host will talk to.
Public NumPlayers As Integer

'These variables are used to calculate country strengths.
Public AttackStrength As Long
Public DefendStrength As Long
Public WaterAttackStrength As Long
Public WaterDefendStrength As Long

'The following four variables keep count of the number of countries
'in each category, for display purposes when right-clicking.
Public AttLndNum As Long
Public AttWtrNum As Long
Public DefLndNum As Long
Public DefWtrNum As Long

'The following variables control the random events that appear from
'time to time.  CheckedForEvent is false until the probability of an
'event has been calculated.  When EventInProgress is true, it means
'that game flow is suspended until it goes false again.
Public CheckedForEvent As Boolean
Public EventInProgress As Boolean

'TroopMoveSrc needs to be global so that we can keep track of which
'country is currently selected during a troop movement.
Public TroopMoveSrc As Long

'The path to the last map played and its time stamp.
Public LastMapPath As String
Public LastMapStamp As String

'Statistical information storage.
Public STATattacked(MAX_PLAYERS, MAX_PLAYERS) As Long
Public STATovertaken(MAX_PLAYERS, MAX_PLAYERS) As Long
Public STATkilled(MAX_PLAYERS, MAX_PLAYERS) As Long
Public STATdefeated(MAX_PLAYERS, MAX_PLAYERS) As Integer
Public STATscore(MAX_PLAYERS) As Single
Public STATrank(MAX_PLAYERS) As Long
Public STATcountries(MAX_PLAYERS) As Long
Public STATtroops(MAX_PLAYERS) As Long

'High score information storage.
Public Const NUM_HI_SCORES As Integer = 10
Public HiScore(NUM_HI_SCORES) As Long           'The 10 high scores for this map.
Public HiScoreName(NUM_HI_SCORES) As String     'And the people who got them!
Public HiScoreColor(NUM_HI_SCORES) As Integer   'The color of each.

'The following is a bunch of declares for sound.
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ALIAS          As Long = &H10000    '  name is a WIN.INI [sounds] entry
Public Const SND_ALIAS_ID       As Long = &H110000   '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START    As Long = 0          '  must be > 4096 to keep strings in same section of resource file
Public Const SND_APPLICATION    As Long = &H80       '  look for application specific association
Public Const SND_ASYNC          As Long = &H1        '  play asynchronously
Public Const SND_FILENAME       As Long = &H20000    '  name is a file name
Public Const SND_LOOP           As Long = &H8        '  loop the sound until next sndPlaySound
Public Const SND_MEMORY         As Long = &H4        '  lpszSoundName points to a memory file
Public Const SND_NODEFAULT      As Long = &H2        '  silence not default, if sound not found
Public Const SND_NOSTOP         As Long = &H10       '  don't stop any currently playing sound
Public Const SND_NOWAIT         As Long = &H2000     '  don't wait if the driver is busy
Public Const SND_PURGE          As Long = &H40       '  purge non-static events for task
Public Const SND_RESERVED       As Long = &HFF000000 '  In particular these flags are reserved
Public Const SND_RESOURCE       As Long = &H40004    '  name is a resource name or atom
Public Const SND_SYNC           As Long = &H0        '  play synchronously (default)
Public Const SND_TYPE_MASK      As Long = &H170007
Public Const SND_VALID          As Long = &H1F       '  valid flags          / ;Internal /
Public Const SND_VALIDFLAGS     As Long = &H17201F   '  Set of valid flag bits.  Anything outside
Public Const SND_GAME_FLAGS     As Long = SND_ASYNC Or SND_NOSTOP Or SND_NOWAIT 'Or SND_PURGE

'Paths.
Public INIpath As String
Public HLPpath As String
Public MAPpath As String
Public SFXpath As String

'flag that the game is ending (ie, the user wants to exit)
Public GameEnding As Boolean


Public Sub SetupPaths()

INIpath = App.Path & "\Fracas.ini"
HLPpath = App.Path & "\Fracas.hlp"
MAPpath = App.Path & "\Maps\"
SFXpath = App.Path & "\Sfx\"

End Sub

Public Function NumNetworkPlayers() As Integer

Dim i As Integer

NumNetworkPlayers = 0

For i = 1 To 6
  If PlayerType(i) = PTYPE_NETWORK Then NumNetworkPlayers = NumNetworkPlayers + 1
Next i

End Function
