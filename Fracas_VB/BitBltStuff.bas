Attribute VB_Name = "BitBlitStuff"
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest

'RollOver tells us that a map exists on the screen and
'we can therefore interact with it.
Public RollOver As Boolean

'These parameters are used to parse menu settings.
Public MinLakeSize As Integer
Public MaxCountrySize As Integer
Public NumCountries As Integer
Public CFGCountrySize As Integer
Public CFGCheckedPct As Integer
Public CFGIslands As Integer
Public CFGLakeSize As Integer
Public CFGCoast As Integer
Public CFGProportion As Integer
Public CFGShape As Integer
Public CFGColor As Integer
Public CFGBorders As Integer
Public CFGInitTroopPl As Integer
Public CFGInitTroopCt As Integer
Public u As Integer
Public LandPct As Double
Public PropPct As Double
Public ShapePct As Double
Public CoastPctKeep As Double
Public IslePctKeep As Double

'These are the screen dimensions.
Public Const WIDE = 656
Public Const TALL = 472

'Make some map variables public.
Public MyMap As Map
Public CurrentMouse As Long

'These are the offsets (in pixels) for the map within the form container.
Public Const XOFFSET = 10
Public Const YOFFSET = 1

'The maximum number of balls that we can track.
Public Const BALLMAX = 500

'The maximum number of waves in the ocean at one time.
Public Const WAVEMAX = 100

'The gravitational constant (in pixels/timer cycle^2)
Public Const GRAVITY = 1

'These constants contain properties for each type of explosion used.
'Note that some of these are multipliers, not absolutes.
'Little splashes in the water.
Public Const SPLASH_COUNT = 15
Public Const SPLASH_INTENSITY = 15
Public Const SPLASH_SPREAD = 20
Public Const SPLASH_ELASTIC = 0
Public Const SPLASH_SIZE = 2
Public Const SPLASH_COLOR = 0
'A small explosion.
Public Const SMALL_COUNT = 0.2
Public Const SMALL_INTENSITY = 0.03
Public Const SMALL_SPREAD = 25
Public Const SMALL_ELASTIC = 50
Public Const SMALL_SIZE = 2
'A medium explosion.
Public Const MED_COUNT = 0.3
Public Const MED_INTENSITY = 0.05
Public Const MED_SPREAD = 30
Public Const MED_ELASTIC = 55
Public Const MED_SIZE = 3
'A big explosion.
Public Const BIG_COUNT = 0.4
Public Const BIG_INTENSITY = 0.07
Public Const BIG_SPREAD = 35
Public Const BIG_ELASTIC = 60
Public Const BIG_SIZE = 4

Public Sub DrawMap()

'These commands draw the map.  MapBuffer is where the map is drawn.
MyMap.DisplayMap Land!LandMap.hdc, Land!MapBuffer.hdc, CFGBorders

End Sub
