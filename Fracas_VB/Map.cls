VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim XDimension As Integer
Dim YDimension As Integer

Dim Direction As Long
Dim LastDirection As Long
Dim TryNumber As Long
Dim OneAway As Long
Dim Country As Long
Dim TempNum As Long

Dim CountryDone As Boolean
Dim FoundLand As Boolean
Dim FilledIn As Boolean

Dim u As Integer

Dim NumCountries As Integer

Public LakeCode As Long
Public MaxCountrySize As Integer
Public MinLakeSize As Integer

Dim CountrySize As Integer
Dim TentNum As Integer
Dim Block As Integer
Dim BlockTry As Integer
Dim FudgeCounter As Integer

Dim MapFileName As String
Dim MapTimeStamp As String  'Used to link maps and saved games.

Dim Map(XDIM1024x768, YDIM1024x768) As Long
'Map Legend:
'0   = Plain water during build process (empty square)
'x   = Country number x
'999 = Coastline during build process
'1000 and up = Bodies of water (determined after land placement)

Dim Tent(2, (MAX_COUNTRY_SIZE * (1 + MAX_PROP_PCT))) As Long
Dim TentOrder((MAX_COUNTRY_SIZE * (1 + MAX_PROP_PCT))) As Long
'Tent(1,a) = x coordinate of point a
'Tent(2,a) = y coordinate of point a
'TentOrder = a random ordering of the elements in Tent.

Dim Seed(XDIM1024x768, YDIM1024x768) As Single
'The seed matrix contains the details of the random parts of the map.
'This is for display purposes only -- Just a tweak.

Dim Neighbor() As Long
'The Neighbor array will be populated with data about which
'countries border others:
'For example:  z = Neighbor(x,y)
'x  = Country in question
'y  = Index number
'z  = The country number of a neighboring country.  Unused indexes return 0.
'There will be no more than MAXIMUM_NEIGHBORS possible borders detected,
'but this should be changed if the resolution is increased.

Dim DisplaySpot() As Long
'The DisplaySpot array will hold the x and y coordinates of a place on
'each country to store information.  This will be populated once the map
'is drawn.
'For example:  z = DisplaySpot(x,y)
'x = Country in question
'If y = 1, z = x coordinate of spot
'If y = 2, z = y coordinate of spot
'If y = 3, z = number of squares of surrounding country (used during build)

Dim Owners() As Integer
'The Owners array holds the owner of each country.
'0 = Neutral territory/unoccupied
'x = Controlled by player x

Dim CountryClr() As Long
'The CountryClr array holds the color of each country.

Dim Troops() As Long
'The Troops array holds the number that will be displayed on top of each country.

Dim Special() As Integer
'The Special array tells whether or not this country has special status.
'These values are binary and can be combined for multiple effects.
'0 = Nothing special about this country
'1 = This country is a PORT (can participate in naval warfare).
'2 = This country is an HQ (the player's starting country).
'4 = ...

Dim CName() As String
'The CName array holds the name of each country, for display purposes.

Dim WName() As String
'The WName array holds the name of a body of water.  Subtract 1000
'from the water ID to get the index.  For example, the name of water
'mass 1005 would be WName(5).

Dim FlashArray() As Integer
'FlashArray holds data regarding how to highlight each country and is
'updated by a timer mechanism in GraphicFX.bas.
'FlashArray(x, 1) = 1 if country x is highlighted currently, 0 if not
'FlashArray(x, 2) = decrementing counter for country x
'FlashArray(x, 3) = timer value for each flash of country x
'FlashArray(x, 4) = number of times left to flash country x

Public Property Get Xsize() As Integer
    Xsize = XDimension
End Property

Public Property Get Ysize() As Integer
    Ysize = YDimension
End Property

Public Property Get NumberOfCountries() As Integer
    NumberOfCountries = NumCountries
End Property

Public Property Get MaxNeighbors() As Integer
    MaxNeighbors = MAXIMUM_NEIGHBORS
End Property

Public Property Get Grid(x As Integer, y As Integer) As Long
    'We put bounds on the inputs so as not to generate subscript out
    'of range errors.  This is because balls may fly one or two
    'squares off of the map before they're caught.  :)
    If x > XDimension Then x = XDimension
    If y > YDimension Then y = YDimension
    If x < 1 Then x = 1
    If y < 1 Then y = 1
    Grid = Map(x, y)
End Property

Public Property Let Grid(x As Integer, y As Integer, z As Long)
    'We put bounds on the inputs so as not to generate subscript out
    'of range errors.  This is because balls may fly one or two
    'squares off of the map before they're caught.  :)
    If x > XDimension Then x = XDimension
    If y > YDimension Then y = YDimension
    If x < 1 Then x = 1
    If y < 1 Then y = 1
    Map(x, y) = z
End Property

Public Property Get Neighbors(x As Long, y As Integer) As Long
    Neighbors = Neighbor(x, y)
End Property

Public Property Get DigitCoords(x As Long, y As Integer) As Long
    DigitCoords = DisplaySpot(x, y)
End Property

Public Property Get Owner(x As Long) As Integer
    Owner = Owners(x)
End Property

Public Property Let Owner(x As Long, y As Integer)
    Owners(x) = y
End Property

Public Property Get CountryColor(x As Long) As Integer
    CountryColor = CountryClr(x)
End Property

Public Property Let CountryColor(x As Long, y As Integer)
    CountryClr(x) = y
End Property

Public Property Get TroopCount(x As Long) As Long
    TroopCount = Troops(x)
End Property

Public Property Let TroopCount(x As Long, y As Long)
    Troops(x) = y
End Property

Public Property Get CountryType(x As Long) As Long
    CountryType = Special(x)
End Property

Public Property Let CountryType(x As Long, y As Long)
    Special(x) = y
End Property

Public Property Get CountryName(xxx As Long) As String
    CountryName = CName(xxx)
End Property

Public Property Let CountryName(xxx As Long, yyy As String)
    CName(xxx) = yyy
End Property

Public Property Get WaterName(xxx As Long) As String
    WaterName = WName(xxx)
End Property

Public Property Let WaterName(xxx As Long, yyy As String)
    WName(xxx) = yyy
End Property

Public Property Get MapName() As String
    MapName = MapFileName
End Property

Public Property Let MapName(x As String)
    MapFileName = x
End Property

Public Property Get MapStamp() As String
    MapStamp = MapTimeStamp
End Property

Public Property Let MapStamp(x As String)
    MapTimeStamp = x
End Property

Public Property Get Flash(CountryID As Long, ParamID As Integer) As Integer
    'We put bounds on the inputs so as not to generate subscript out
    'of range errors.  This is because balls may fly one or two
    'squares off of the map before they're caught.  :)
    If CountryID > NumCountries Then CountryID = NumCountries
    If ParamID > 4 Then ParamID = 4
    If CountryID < 1 Then CountryID = 1
    If ParamID < 1 Then ParamID = 1
    Flash = FlashArray(CountryID, ParamID)
End Property

Public Property Let Flash(CountryID As Long, ParamID As Integer, z As Integer)
    'We put bounds on the inputs so as not to generate subscript out
    'of range errors.  This is because balls may fly one or two
    'squares off of the map before they're caught.  :)
    If CountryID > NumCountries Then CountryID = NumCountries
    If ParamID > 4 Then ParamID = 4
    If CountryID < 1 Then CountryID = 1
    If ParamID < 1 Then ParamID = 1
    FlashArray(CountryID, ParamID) = z
End Property

Public Sub CreateMap(CountryCount As Integer, _
                    MaximumCountrySize As Integer, _
                    MinimumLakeSize As Integer, _
                    LandPct As Double, _
                    PropPct As Double, _
                    ShapePct As Double, _
                    CoastPctKeep As Double, _
                    IslePctKeep As Double)

Dim i As Long
Dim j As Long
Dim x As Long
Dim y As Long
Dim LocalMax As Integer
Dim LocalMin As Integer
Dim StillPossible As Boolean
Dim NewOrder As Boolean

NumCountries = CountryCount

'We want some passed variables to stay public.
MaxCountrySize = MaximumCountrySize
MinLakeSize = MinimumLakeSize

'Redimension our arrays to save some memory.
ReDim Neighbor(NumCountries, MAXIMUM_NEIGHBORS)
ReDim CountryClr(NumCountries)
ReDim Owners(NumCountries)
ReDim DisplaySpot(NumCountries, 3)
ReDim Troops(NumCountries)
ReDim Special(NumCountries)
ReDim CName(NumCountries)
ReDim FlashArray(NumCountries, 4)

Do   'Loop through everything until we pass spec.

  'All of map is considered coastline for placing of first country.
  For i = 1 To XDimension
    For j = 1 To YDimension
      Map(i, j) = TILEVAL_COASTLINE
    Next j
  Next i
  
  Country = 0
  
  Do While Country < NumCountries  'Main Country Loop.
    Country = Country + 1
    StillKicking   'Update our commentary line.
      
    FudgeCounter = 0
    
    CountryDone = False
    Do Until CountryDone    'This Country Loop.
    
      'First, make sure that we *can* place a country on this map.
      'if we have a coastline or 0 we can still do it.
      StillPossible = False
      For i = 1 To XDimension
        For j = 1 To YDimension
          If (Map(i, j) = 0) Or (Map(i, j) = TILEVAL_COASTLINE) Then
            StillPossible = True
            Exit For
          End If
        Next j
      Next i
      StillKicking
'      DoEvents
      If StillPossible Then
        'Find the starting position of the next country.
        Done = False
        Do Until Done = True   'Place loop.
          x = Int(Rnd(1) * XDimension) + 1
          y = Int(Rnd(1) * YDimension) + 1
          'Take care of the None or Lots of Islands options.
          If ((CFGIslands = 1) And (Map(x, y) = TILEVAL_COASTLINE)) Then
            Done = True
          End If
          If (CFGIslands = 3) And ((Map(x, y) = 0) Or (Map(x, y) = TILEVAL_COASTLINE)) Then
            Done = True
          End If
          'The 'Some Islands' option is more complex.
          'If we found coastline, there's a large chance of keeping it.
          'If we found empty, there's a small chance of keeping it.
          If CFGIslands = 2 Then
            If ((Map(x, y) = TILEVAL_COASTLINE) And (Rnd(1) < CoastPctKeep)) Or ((Map(x, y) = 0) And (Rnd(1) < IslePctKeep)) Then
              Done = True
            End If
          End If
        Loop   'Place loop.
      Else
        'Couldn't find a place for it, so we're done.
        EndItAll
      End If
      
      If StillPossible Then
         
        'Clear out the last country by clearing out the tentative array.
        For i = 1 To 2
          For j = 1 To (MAX_COUNTRY_SIZE * (1 + MAX_PROP_PCT))
            Tent(i, j) = 0
          Next j
        Next i
        'This is our first tentative block.
        Tent(1, 1) = x
        Tent(2, 1) = y
        TentNum = 1
              
        'Get a random size for this country.
        If CFGCountrySize <> 6 Then
          If Rnd(1) < 0.5 Then
            'Subtract a random amount within the proportional limit.
            CountrySize = MaxCountrySize - Int(Rnd(1) * MaxCountrySize * PropPct)
          Else
            'Add a random amount within the proportional limit.  This will keep
            'all countries equal to MaxCountrySize on average.
            '(So it's not really a MAX then, is it?)
            CountrySize = MaxCountrySize + Int(Rnd(1) * MaxCountrySize * PropPct)
          End If
        Else
          'Hodge-Podge:  Pick anything between min and max.
          LocalMax = (MAX_COUNTRY_SIZE * (1 + PropPct))
          LocalMin = (MIN_COUNTRY_SIZE * (1 - PropPct))
          CountrySize = Int(Rnd(1) * (LocalMax - LocalMin)) + LocalMin
        End If
                   
        NewOrder = True
        'Now we search for blocks contiguous to this one.
        Do Until (TentNum = CountrySize)    'Block Search Loop.
        
          'Each block must be tried only once. Order them randomly,
          'and pick a first choice before we get into this DO.  Then jump
          'to the next one in order for each iteration.  The first choice
          'will be based on the irregularity of the country.
    
          If NewOrder = True Then
            'Set up our array of integers.  There are TentNum integers to mix up.
            'This is the order that we check elements in the Tent array.
            For i = 1 To TentNum
              TentOrder(i) = i
            Next i
            'Now randomize them a bit.  Don't move the last block.
            If TentNum >= 4 Then
              For i = 1 To TentNum * 2
                Do
                  x = Int(Rnd(1) * (TentNum - 1)) + 1
                  y = Int(Rnd(1) * (TentNum - 1)) + 1
                Loop Until (x <> y)
                Block = TentOrder(x)
                TentOrder(x) = TentOrder(y)
                TentOrder(y) = Block
              Next i
            End If
            
            'When BlockTry reaches TentNum, we've tried to bud off of every
            'square in this country.
            BlockTry = 0
           
            'Regular or Irregular?
            'Block is which element in the Tent array we're trying to 'bud' off of.
            If Rnd(1) < ShapePct Then
              'Start at the first random block.
              Block = 1
            Else
              'Start at the last block we used.  Makes it more irregular.
              Block = TentNum
            End If
          End If   'NewOrder = true
        
          'Get the coordinates of a block in this country.
          x = Tent(1, TentOrder(Block))
          y = Tent(2, TentOrder(Block))
          'Pick a random direction for a contiguous block.
          Call NewDirection
          Done = False
          Do Until (TryNumber = 5 Or Done = True)   'Each dir.
            Select Case Direction
              Case UP
                OneAway = y - 1
                If OneAway > 0 Then
                  'Check and see what's up.
                  If ClearDirection(x, OneAway) = True Then
                    'Map is clear in that direction.
                    y = y - 1
                    Call WriteToTent(x, y)
                    Done = True
                  Else
                    'Blocked.  Try new direction.
                    Call NextDirection
                  End If
                Else
                  'We went off the top.  Try new direction.
                  Call NextDirection
                End If
              Case DOWN
                OneAway = y + 1
                If OneAway <= YDimension Then
                  'Check and see what's down.
                  If ClearDirection(x, OneAway) = True Then
                    'Map is clear in that direction.
                    y = y + 1
                    Call WriteToTent(x, y)
                    Done = True
                  Else
                    'Blocked.  Try new direction.
                    Call NextDirection
                  End If
                Else
                  'We went off the bottom.  Try new direction.
                  Call NextDirection
                End If
              Case LEFTY
                OneAway = x - 1
                If OneAway > 0 Then
                  'Check and see what's up.
                  If ClearDirection(OneAway, y) = True Then
                    'Map is clear in that direction.
                    x = x - 1
                    Call WriteToTent(x, y)
                    Done = True
                  Else
                    'Blocked.  Try new direction.
                    Call NextDirection
                  End If
                Else
                  'We went off the left.  Try new direction.
                  Call NextDirection
                End If
              Case RIGHTY
                OneAway = x + 1
                If OneAway <= XDimension Then
                  'Check and see what's up.
                  If ClearDirection(OneAway, y) = True Then
                    'Map is clear in that direction.
                    x = x + 1
                    Call WriteToTent(x, y)
                    Done = True
                  Else
                    'Blocked.  Try new direction.
                    Call NextDirection
                  End If
                Else
                  'We went off the right.  Try new direction.
                  Call NextDirection
                End If
            End Select
          Loop   'Each dir.
          If TryNumber = 5 Then
            'This block is boxed in.  Try the next block in the country.
            NewOrder = False
            Block = Block + 1
            If Block > TentNum Then
              Block = 1
            End If
            'Have we tried all blocks?
            BlockTry = BlockTry + 1
            If BlockTry > TentNum Then
              'This Country cannot fit.  Need new starting location.
              'Remove this small area from further consideration -- probably a lake.
              For i = 1 To TentNum
                'This space won't get picked next time.
                Map(Tent(1, i), Tent(2, i)) = UNUSABLE_GRID
              Next i
              TentNum = CountrySize   'Fudging out of loop.
              FudgeCounter = FudgeCounter + 1
              StillPossible = False
              'Move on to next country if we can't place this one
              'within a reasonable number of tries.
              If (FudgeCounter > MAX_COUNTRY_TRIES) Then
                'Screw it ... taking too long.  We're done.
                EndItAll
              End If
            End If
          Else
            NewOrder = True
          End If
        Loop     'Block Search Loop.
      
        'Yay!  Our country is sitting in the Tent array.
        'Let's copy it over to the Map.
        'If we blew it last time, then we don't need to do anything here.
        If StillPossible Then
          For i = 1 To CountrySize
            Map(Tent(1, i), Tent(2, i)) = Country
          Next i
          'Let's outline all countries on the map with coastline.
          'Because if there are no islands, we must build on coastline next time.
          OutlineCoastline
        End If   'StillPossible.
      End If   'StillPossible.
    Loop     'This Country Loop.
  Loop     'Main country loop.
  
  'Give us our final coastline.
  Call OutlineCoastline
  
  'Fill in lakes according to lakesize parameter.
  Call FillLakes

'Populate the Neighbors array.
Loop Until FindNeighbors

'Find the best display spot for each country.
Call FindSpots

'Initialize the seed matrix.
Call BuildSeedMatrix

'Clear out the troops, special, and flash arrays.  Name each country too.
For i = 1 To NumCountries
  Troops(i) = 0
  Special(i) = 0
  For j = 1 To 4
    FlashArray(i, j) = 0
  Next j
  If Neighbor(i, 2) = 0 And Neighbor(i, 1) > TILEVAL_COASTLINE Then
    'We have an island!  Let's give it an island name.
    CName(i) = GenerateName(2, 6, False)
    TempNum = Int(Rnd(1) * 3) + 23
    If TempNum = 25 Then
      CName(i) = NameSpice(TempNum) & " " & CName(i)
    Else
      CName(i) = CName(i) & " " & NameSpice(TempNum)
    End If
  Else
    CName(i) = GenerateName(2, 6, True)
  End If
Next i

'Now let's name each body of water.
If LakeCode = 1001 Then
  'There were no bodies of water in this map.
  Exit Sub
End If

'Let's set up the array.
ReDim WName(LakeCode - 1000)

For i = (LakeCode - 1000) To 1 Step -1
  'Search the borders of the map for this lakecode.
  'Check the top and bottom borders.
  Done = False
  For j = 1 To XDimension
    If Map(j, 1) = (i + 1000) Or Map(j, YDimension) = (i + 1000) Then
      'Yup, it's on the border.
      Done = True
      Exit For
    End If
  Next j
  'Check the left and right borders.
  For j = 1 To YDimension
    If Map(1, j) = (i + 1000) Or Map(XDimension, j) = (i + 1000) Then
      'Yup, it's on the border.
      Done = True
      Exit For
    End If
  Next j
  If Done = False Then
    'We haven't found it on the edge, so it must be a LAKE.
    Select Case Int(Rnd(1) * 11) + 1
    Case 1, 2, 3:
      WName(i) = GenerateName(2, 6, False) & " Lake"
    Case 4, 5, 6:
      WName(i) = "Lake " & GenerateName(2, 6, False)
    Case 7:
      WName(i) = GenerateName(2, 6, False) & " Vista"
    Case 8, 9:
      WName(i) = GenerateName(2, 6, False) & " Sea"
    Case 10, 11:
      WName(i) = "The Sea of " & GenerateName(2, 6, False)
    End Select
  Else
    'We found it on the edge of the map, so this is an OCEAN.
    Select Case Int(Rnd(1) * 7) + 1
    Case 1:
      WName(i) = GenerateName(2, 6, False) & " Ocean"
    Case 2, 3:
      WName(i) = GenerateName(2, 6, False) & " Harbor"
    Case 4, 5:
      WName(i) = GenerateName(2, 6, False) & " Bay"
    Case 6, 7:
      WName(i) = "The Bay of " & GenerateName(2, 6, False)
    End Select
  End If
Next i

'Finally, this is a fresh map so it hasn't been saved yet.
MapFileName = vbNullString
MapTimeStamp = vbNullString

End Sub

Private Function WriteToTent(x As Long, y As Long)

'Found a good contiguous block, lets record its X and Y.
TentNum = TentNum + 1
Tent(1, TentNum) = x
Tent(2, TentNum) = y

'Was this the last tentative block?  If so, we need to signal we're done.
If TentNum = CountrySize Then
  CountryDone = True
End If

End Function

Private Function ClearDirection(Xcheck, Ycheck) As Boolean

Dim i As Long

'This function checks the Map and Tent arrays to see if
'the suggested block is already used.
ClearDirection = True

'Is the suggested block part of another country?
If Map(Xcheck, Ycheck) > 0 And Map(Xcheck, Ycheck) < TILEVAL_COASTLINE Then
  ClearDirection = False
End If

'Is the suggested block part of the current country (or lake)?
For i = 1 To TentNum
  If Tent(1, i) = Xcheck And Tent(2, i) = Ycheck Then
    ClearDirection = False
  End If
Next i

End Function

Private Sub OutlineCoastline()

Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer

For i = 1 To XDimension
  For j = 1 To YDimension
    'If this was the first country, we erase all coastline.
    'Also erase coastline if we find our temporary block-out code.
    If ((Country = 1) And (Map(i, j) = TILEVAL_COASTLINE)) Or (Map(i, j) = UNUSABLE_GRID) Then
      Map(i, j) = 0
    End If
    If Map(i, j) = 0 Or Map(i, j) > TILEVAL_COASTLINE Then
      FoundLand = False
      'Land up?
      x = i
      y = j - 1
      If y > 0 Then
        If Map(x, y) > 0 And Map(x, y) < TILEVAL_COASTLINE Then FoundLand = True
      End If
      'Land down?
      x = i
      y = j + 1
      If y <= YDimension Then
        If Map(x, y) > 0 And Map(x, y) < TILEVAL_COASTLINE Then FoundLand = True
      End If
      'Land left?
      x = i - 1
      y = j
      If x > 0 Then
        If Map(x, y) > 0 And Map(x, y) < TILEVAL_COASTLINE Then FoundLand = True
      End If
      'Land right?
      x = i + 1
      y = j
      If x <= XDimension Then
        If Map(x, y) > 0 And Map(x, y) < TILEVAL_COASTLINE Then FoundLand = True
      End If
      'Place coastline.
      If FoundLand = True Then
        Map(i, j) = TILEVAL_COASTLINE
      Else
        Map(i, j) = 0
      End If
    End If
  Next j
Next i

End Sub

Private Sub EndItAll()

'We call this when we can go no further with placing countries.

CountryDone = True
NumCountries = Country - 1
'Redimension our arrays to save some memory.
ReDim Neighbor(NumCountries, MAXIMUM_NEIGHBORS)
ReDim CountryClr(NumCountries)
ReDim Owners(NumCountries)
ReDim DisplaySpot(NumCountries, 3)
ReDim Troops(NumCountries)
ReDim Special(NumCountries)
ReDim CName(NumCountries)
ReDim FlashArray(NumCountries, 4)

End Sub

Private Function GetAdjacentColor(Xcheck, Ycheck)

If Map(Xcheck, Ycheck) <> 0 And Map(Xcheck, Ycheck) <> TILEVAL_COASTLINE And Country = 0 Then
  Country = Map(Xcheck, Ycheck)
End If

End Function

Private Function IsOnCoastLine(ii As Long, jj As Long) As Boolean

'This sub returns true if the passed coordinate is next to coastline.
If LookRight(ii, jj) >= TILEVAL_COASTLINE Or LookLeft(ii, jj) >= TILEVAL_COASTLINE Or _
        LookUp(ii, jj) >= TILEVAL_COASTLINE Or LookDown(ii, jj) >= TILEVAL_COASTLINE Or _
        LookRight(ii, jj) = 0 Or LookLeft(ii, jj) = 0 Or _
        LookUp(ii, jj) = 0 Or LookDown(ii, jj) = 0 Then
  IsOnCoastLine = True
Else
  IsOnCoastLine = False
End If


End Function

Private Function LookRight(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookRight = -1
xx = ii + 1
yy = jj
If xx <= XDimension Then
  LookRight = Map(xx, yy)
End If

End Function

Private Function LookLeft(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookLeft = -1
xx = ii - 1
yy = jj
If xx > 0 Then
  LookLeft = Map(xx, yy)
End If

End Function
      
Private Function LookDown(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookDown = -1
xx = ii
yy = jj + 1
If yy <= YDimension Then
  LookDown = Map(xx, yy)
End If
      
End Function

Private Function LookUp(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookUp = -1
xx = ii
yy = jj - 1
If yy > 0 Then
  LookUp = Map(xx, yy)
End If
      
End Function

Private Function LookUpRight(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookUpRight = -1
xx = ii + 1
yy = jj - 1
If (yy > 0) And (xx <= XDimension) Then
  LookUpRight = Map(xx, yy)
End If
      
End Function

Private Function LookUpLeft(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookUpLeft = -1
xx = ii - 1
yy = jj - 1
If (yy > 0) And (xx > 0) Then
  LookUpLeft = Map(xx, yy)
End If
      
End Function

Private Function LookDownRight(ii As Long, jj As Long) As Long

Dim xx As Long
Dim yy As Long

LookDownRight = -1
xx = ii + 1
yy = jj + 1
If (yy <= YDimension) And (xx <= XDimension) Then
  LookDownRight = Map(xx, yy)
End If
      
End Function

Private Function NextDirection()

'The last direction didn't work, so we try the next until we've done them all.
Direction = Direction + 1
If Direction = 5 Then Direction = 1
TryNumber = TryNumber + 1

End Function

Private Function NewDirection()

'Record which way we went last time.
LastDirection = Direction
If LastDirection < 1 Or LastDirection > 4 Then
  LastDirection = RandomDir
End If
Direction = RandomDir
'This will be our first try for a new way to go.
TryNumber = 1

End Function

Public Function DisplayMap(Source As Long, dest As Long)

Dim PieceNum As Integer

Dim Xadj As Integer
Dim Yadj As Integer

Dim UpPiece As Long
Dim DownPiece As Long
Dim LeftPiece As Long
Dim RightPiece As Long
Dim UpRightPiece As Long
Dim UpLeftPiece As Long
Dim DownRightPiece As Long

Dim ii As Long
Dim jj As Long
Dim xx As Long
Dim yy As Long
Dim i As Long
Dim ThisPiece As Long
Dim PieceXcoord As Integer
Dim PieceXbase As Integer
Dim LineLength As Long
Dim MySeedVal As Long

'Let's update the map for everyone to see.
For ii = 1 To XDimension
  For jj = 1 To YDimension
    
    ThisPiece = Map(ii, jj)
    If ThisPiece > 0 And ThisPiece < TILEVAL_COASTLINE Then
      
      'This routine will determine which 'piece' gets used for this block.
      'Initialize our counters.
      PieceNum = 0
      UpPiece = LookUp(ii, jj)
      DownPiece = LookDown(ii, jj)
      RightPiece = LookRight(ii, jj)
      LeftPiece = LookLeft(ii, jj)
      
      'See if there's water or coastline to the right of us.
      If RightPiece = 0 Or RightPiece >= TILEVAL_COASTLINE Then PieceNum = PieceNum + 1   'Add a binary 1
      
      'See if there's water or coastline to the left of us.
      If LeftPiece = 0 Or LeftPiece >= TILEVAL_COASTLINE Then PieceNum = PieceNum + 2   'Add a binary 2
      
      'See if there's water or coastline below us.
      If DownPiece = 0 Or DownPiece >= TILEVAL_COASTLINE Then PieceNum = PieceNum + 4   'Add a binary 4
      
      'See if there's water or coastline above us.
      If UpPiece = 0 Or UpPiece >= TILEVAL_COASTLINE Then PieceNum = PieceNum + 8  'Add a binary 8
      
      'At this point, we know what shape this piece should be.

      'Calculate the x coordinates in LandMap, since we'll use it a lot.
      PieceXbase = GFX_NUM_LAND_PIECES * Int(Seed(ii, jj) * GFX_NUM_LAND_PIECE_SETS)
      PieceXcoord = PieceNum + PieceXbase
      
      'Draw the land piece in.
      Call DrawBlock(Source, dest, ii, jj, CorrectColor(ThisPiece), PieceXcoord)
            
      'Draw the correct base border if this piece is different than ones around it.
      'Only do up and right, since that's where the stripes are drawn.
      'Check up.
      If CFGBorders = 5 Then
        If (UpPiece <> ThisPiece) And (UpPiece > 0) And (UpPiece < TILEVAL_COASTLINE) Then
          Call DrawLine(Source, dest, ii, jj, DIR_UP)
        End If
        'Check right.
        If (RightPiece <> ThisPiece) And (RightPiece > 0) And (RightPiece < TILEVAL_COASTLINE) Then
          Call DrawLine(Source, dest, ii, jj, DIR_RIGHT)
        End If
      End If
            
      'If this country is an HQ, we need to draw the HQ symbol on it...
      If (Special(ThisPiece) And 2) = 2 Then
        'First, copy the HQ logo and its mask to the working area.
        u% = BitBlt(Source, GFX_WORK_X_COORD, GFX_HQ_Y_COORD, GFX_GRID, GFX_GRID_X2, Source, GFX_HQ_X_COORD, GFX_HQ_Y_COORD, SRCCOPY)
        'Now, add the mask for the appropriate land piece in on top of the HQ mask.
        u% = BitBlt(Source, GFX_WORK_X_COORD, GFX_WHITE_Y_COORD, GFX_GRID, GFX_GRID, Source, (PieceXcoord * GFX_GRID), GFX_LAND_MASK_Y_COORD, SRCPAINT)
        'Invert the land piece mask by XORing it into whiteness.
        'There is a permanent white square in landmap at 17,14.
        'First copy the white square.
        u% = BitBlt(Source, GFX_WHITE_X_COORD, GFX_HQ_Y_COORD, GFX_GRID, GFX_GRID, Source, GFX_WHITE_X_COORD, GFX_WHITE_Y_COORD, SRCCOPY)
        'Then XOR the mask in.
        u% = BitBlt(Source, GFX_WHITE_X_COORD, GFX_HQ_Y_COORD, GFX_GRID, GFX_GRID, Source, (PieceXcoord * GFX_GRID), GFX_LAND_MASK_Y_COORD, SRCINVERT)
        'Add the inverted mask to the HQ logo to blank out parts we don't want to draw.
        u% = BitBlt(Source, GFX_WORK_X_COORD, GFX_HQ_Y_COORD, GFX_GRID, GFX_GRID, Source, GFX_WHITE_X_COORD, GFX_HQ_Y_COORD, SRCAND)
        'Now draw it!
        Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, GFX_NUM_LAND_PIECES - 1)
      End If
      
      'Now we check for the edit pieces between countries to smooth out borders.
      'Note that we *only* write on the current square.  We don't butt into
      'someone else's country!
      
      'All this does is smooth out a corner of our square if a country is diagonal
      'to us.
      
      'Remember that borders are drawn on the top edge and right edge of each tile.
      
      'Check upper left.
      If (UpPiece = LeftPiece) And (UpPiece <> ThisPiece) And (UpPiece > 0) And (UpPiece < TILEVAL_COASTLINE) Then
        'Now fill in our corner...
        Call DrawCorner(Source, dest, ii, jj, CorrectColor(UpPiece), GFX_EDIT_PIECE_X_BASE + PieceXbase)
      End If
      'Check upper right.
      If (RightPiece = UpPiece) And (RightPiece <> ThisPiece) And (RightPiece > 0) And (RightPiece < TILEVAL_COASTLINE) Then
        Call DrawCorner(Source, dest, ii, jj, CorrectColor(RightPiece), GFX_EDIT_PIECE_X_BASE + 1 + PieceXbase)
      End If
      'Check lower left.
      If (LeftPiece = DownPiece) And (LeftPiece <> ThisPiece) And (LeftPiece > 0) And (LeftPiece < TILEVAL_COASTLINE) Then
        Call DrawCorner(Source, dest, ii, jj, CorrectColor(LeftPiece), GFX_EDIT_PIECE_X_BASE + 2 + PieceXbase)
      End If
      'Check lower right.
      If (DownPiece = RightPiece) And (DownPiece <> ThisPiece) And (DownPiece > 0) And (DownPiece < TILEVAL_COASTLINE) Then
        Call DrawCorner(Source, dest, ii, jj, CorrectColor(DownPiece), GFX_EDIT_PIECE_X_BASE + 3 + PieceXbase)
      End If
      
      If CFGBorders = 5 Then
        'Finally, clean up borders to the top and right if we need to.
        UpRightPiece = LookUpRight(ii, jj)
        UpLeftPiece = LookUpLeft(ii, jj)
        DownRightPiece = LookDownRight(ii, jj)
        MySeedVal = Int(Seed(ii, jj) * GFX_NUM_LAND_PIECE_SETS) + 3
                
        If (ThisPiece <> UpPiece) And (UpPiece < TILEVAL_COASTLINE) And _
                    (ThisPiece = RightPiece) And (ThisPiece = UpRightPiece) Then
          LineLength = 5 - Int(Seed(ii, jj - 1) * GFX_NUM_LAND_PIECE_SETS)
          If MySeedVal < LineLength Then LineLength = MySeedVal
          DrawFixLine Source, dest, ii, jj, CorrectColor(ThisPiece), 3, LineLength
        End If
        If (ThisPiece <> UpPiece) And (UpPiece < TILEVAL_COASTLINE) And _
                    (ThisPiece = LeftPiece) And (ThisPiece = UpLeftPiece) Then
          LineLength = 5 - Int(Seed(ii, jj - 1) * GFX_NUM_LAND_PIECE_SETS)
          If MySeedVal < LineLength Then LineLength = MySeedVal
          DrawFixLine Source, dest, ii, jj, CorrectColor(ThisPiece), 1, LineLength
        End If
        If (ThisPiece <> RightPiece) And (RightPiece < TILEVAL_COASTLINE) And _
                    (ThisPiece = UpPiece) And (ThisPiece = UpRightPiece) Then
          LineLength = 5 - Int(Seed(ii + 1, jj) * GFX_NUM_LAND_PIECE_SETS)
          If MySeedVal < LineLength Then LineLength = MySeedVal
          DrawFixLine Source, dest, ii, jj, CorrectColor(ThisPiece), 4, LineLength
        End If
        If (ThisPiece <> RightPiece) And (RightPiece < TILEVAL_COASTLINE) And _
                    (ThisPiece = DownPiece) And (ThisPiece = DownRightPiece) Then
          LineLength = 5 - Int(Seed(ii + 1, jj) * GFX_NUM_LAND_PIECE_SETS)
          If MySeedVal < LineLength Then LineLength = MySeedVal
          DrawFixLine Source, dest, ii, jj, CorrectColor(ThisPiece), 2, LineLength
        End If
      End If
      
      'Now we will add the cute little piers and docks if this country
      'has a port.
      If (Special(ThisPiece) And 1) = 1 Then
        'This country has a port.  Let's draw in the port piece if it borders
        'water.  In other words, PieceNum is between 1 and 14.
        If PieceNum >= 1 And PieceNum <= 14 Then
          'Yup, this one needs a pier.  Let's calculate the x and y adjustments...
          Xadj = 0
          Yadj = 0
          If (PieceNum = 1) Then Xadj = 3
          If (PieceNum = 2) Then Xadj = -3
          If (PieceNum = 4) Then Yadj = 3
          If (PieceNum = 8) Then Yadj = -3
          If (PieceNum = 5) Then
            Xadj = 2
            Yadj = 2
          End If
          If (PieceNum = 6) Then
            Xadj = -2
            Yadj = 2
          End If
          If (PieceNum = 9) Then
            Xadj = 2
            Yadj = -2
          End If
          If (PieceNum = 10) Then
            Xadj = -2
            Yadj = -2
          End If
          
          Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, _
              ((GFX_NUM_PIER_PIECES * Int(Seed(ii, jj) * GFX_NUM_PIER_PIECE_SETS)) + PieceNum + GFX_PIER_X_BASE), Xadj, Yadj)
        End If
      End If
      
      'We will also add the selected border if two adjacent squares are
      'different countries.  If no border selected, then move on.
      
      If CFGBorders < 5 Then
        'Check up.
        If (UpPiece <> ThisPiece) And (UpPiece > 0) And (UpPiece < TILEVAL_COASTLINE) Then
          Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, 0 + (CFGBorders - 1) * 4)
        End If
        'Check right.
        If (RightPiece <> ThisPiece) And (RightPiece > 0) And (RightPiece < TILEVAL_COASTLINE) Then
          Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, 2 + (CFGBorders - 1) * 4)
        End If
        'Check left.
        If (LeftPiece <> ThisPiece) And (LeftPiece > 0) And (LeftPiece < TILEVAL_COASTLINE) Then
          Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, 3 + (CFGBorders - 1) * 4)
        End If
        'Check down.
        If (DownPiece <> ThisPiece) And (DownPiece > 0) And (DownPiece < TILEVAL_COASTLINE) Then
          Call DrawBlock(Source, dest, ii, jj, GFX_SUPP_Y_LINE, 1 + (CFGBorders - 1) * 4)
        End If
      End If
    End If
  Next jj
Next ii

'Now we will superimpose each country's troop count on top of each country in
'the predetermined display spot.
If (GameMode = GM_GAME_ACTIVE) Or (GameMode = GM_DIALOG_OPEN) Then
                                          'Otherwise we draw numbers on the title screen!
  For i = 1 To NumCountries
    Call DrawNumber(Source, dest, DisplaySpot(i, 1), DisplaySpot(i, 2), Troops(i))
  Next i
End If

End Function

Private Function DrawBlock(Source As Long, dest As Long, x As Long, y As Long, Color As Long, PieceNum As Integer, Optional Xadjust As Integer, Optional Yadjust As Integer)

Dim PieceX As Integer
Dim ColorY As Integer
Dim MaskColor As Integer

'Set up the mask and coords.
If Color > 12 Then
  'We're drawing borders or numbers or something else.
  MaskColor = (Color + 1) * GFX_GRID
  PieceX = GFX_GRID * PieceNum
  ColorY = GFX_GRID * Color
Else
  'We're drawing land.
  MaskColor = GFX_MASK_Y_LINE * GFX_GRID
  PieceX = GFX_GRID * PieceNum
  ColorY = GFX_GRID * (Color - 1)
End If

u% = BitBlt(dest, ((x - 1) * GFX_GRID) + Xadjust, ((y - 1) * GFX_GRID) + Yadjust, GFX_GRID, GFX_GRID, Source, PieceX, MaskColor, SRCAND)
u% = BitBlt(dest, ((x - 1) * GFX_GRID) + Xadjust, ((y - 1) * GFX_GRID) + Yadjust, GFX_GRID, GFX_GRID, Source, PieceX, ColorY, SRCINVERT)

End Function

Private Function DrawCorner(Source As Long, dest As Long, x As Long, y As Long, Color As Long, PieceNum As Integer)

Dim PieceX As Integer
Dim ColorY As Integer
Dim MaskColor As Integer

'Set up the mask and coords.
PieceX = GFX_GRID * PieceNum
ColorY = GFX_GRID * (Color - 1)
If CFGBorders = 5 Then
  'We're drawing borders or numbers or something else.
  MaskColor = (GFX_MASK_Y_LINE + 1) * GFX_GRID
Else
  'We're drawing normal land.
  MaskColor = GFX_MASK_Y_LINE * GFX_GRID
End If

u% = BitBlt(dest, ((x - 1) * GFX_GRID), ((y - 1) * GFX_GRID), GFX_GRID, GFX_GRID, Source, PieceX, MaskColor, SRCAND)
u% = BitBlt(dest, ((x - 1) * GFX_GRID), ((y - 1) * GFX_GRID), GFX_GRID, GFX_GRID, Source, PieceX, ColorY, SRCINVERT)

End Function

Private Function DrawLine(Source As Long, dest As Long, x As Long, y As Long, Direction As Long)

Dim PieceX As Integer
Dim ColorY As Integer
Dim MaskColor As Integer

'Set up the mask and coords.
PieceX = GFX_GRID * (Direction + 1)
ColorY = GFX_GRID * (GFX_MASK_Y_LINE + 1)
MaskColor = ColorY

u% = BitBlt(dest, ((x - 1) * GFX_GRID), ((y - 1) * GFX_GRID), GFX_GRID, GFX_GRID, Source, PieceX, MaskColor, SRCAND)
u% = BitBlt(dest, ((x - 1) * GFX_GRID), ((y - 1) * GFX_GRID), GFX_GRID, GFX_GRID, Source, 0, ColorY, SRCINVERT)

End Function

Private Function DrawFixLine(Source As Long, dest As Long, x As Long, y As Long, Color As Integer, Shape As Integer, Length As Long)

Dim ColorY As Integer

'Set up the mask and coords.
ColorY = GFX_GRID * (Color - 1)

Select Case Shape
  Case 1:
    u% = BitBlt(dest, ((x - 1) * GFX_GRID), ((y - 1) * GFX_GRID), Length, 1, Source, 0, ColorY, SRCCOPY)
  Case 2:
    u% = BitBlt(dest, ((x - 1) * GFX_GRID) + 7, ((y - 1) * GFX_GRID) + (8 - Length), 1, Length, Source, 0, ColorY, SRCCOPY)
  Case 3:
    u% = BitBlt(dest, ((x - 1) * GFX_GRID) + (8 - Length), ((y - 1) * GFX_GRID), Length, 1, Source, 0, ColorY, SRCCOPY)
  Case 4:
    u% = BitBlt(dest, ((x - 1) * GFX_GRID) + 7, ((y - 1) * GFX_GRID), 1, Length, Source, 0, ColorY, SRCCOPY)
End Select

End Function

Private Function RandomDir()

'This function generates a random integer, 1-4.
RandomDir = Int(Rnd(1) * 4) + 1

End Function

Private Function FillLakes()

'This function checks for lakes on the entire map.
'If a lake is under the minimum size, then we fill it in
'with the color of a random adjacent country.
'If a lake is >= the minimum size, we will fill it in with
'the code for the next body of water (1001 and up).
'This will speed up the fill-in procedure so that each body
'of water is only tested once.  Also, we will be able to
'distinguish between lakes when we are done!

Dim i As Long
Dim j As Long
Dim k As Long
Dim x As Long
Dim y As Long

LakeCode = 1001

If MinLakeSize = 0 Then
  'If we chose No Lake Correction, we need to plant seeds so that we
  'can still identify bodies of water.
  For j = 1 To XDimension
    For k = 1 To YDimension
      If Map(j, k) = TILEVAL_COASTLINE Then
        Map(j, k) = LakeCode
        Call LabelBodyOfWater(LakeCode)
        LakeCode = LakeCode + 1
      End If
    Next k
  Next j
  Exit Function
End If

For j = 1 To XDimension
  For k = 1 To YDimension
    'We always start our lake search on coastline.
    FilledIn = False
    If Map(j, k) = TILEVAL_COASTLINE Then
      
      'Clear out the last lake by clearing out the tentative array.
      For i = 1 To MinLakeSize
        Tent(1, i) = 0
        Tent(2, i) = 0
      Next i
      
      Tent(1, 1) = j
      Tent(2, 1) = k
      TentNum = 1
      Country = 0
      Block = 1

      'Now we search for blocks contiguous to this one.
      Do Until TentNum = MinLakeSize
        'Get the coordinates of a block in this country.
        x = Tent(1, Block)
        y = Tent(2, Block)
        'Pick a direction to look for a contiguous block.
        Direction = 1
        Done = False
        Do Until (Direction = 5 Or Done = True)
        Select Case Direction
          Case UP
            OneAway = y - 1
            If OneAway > 0 Then
              'Check and see what's up.
              If ClearDirection(x, OneAway) = True Then
                'Map is clear in that direction.
                y = y - 1
                Call WriteToTent(x, y)
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(x, OneAway)
                Direction = Direction + 1
              End If
            Else
              'We went off the top.  Try new direction.
              Direction = Direction + 1
            End If
          Case DOWN
            OneAway = y + 1
            If OneAway <= YDimension Then
              'Check and see what's down.
              If ClearDirection(x, OneAway) = True Then
                'Map is clear in that direction.
                y = y + 1
                Call WriteToTent(x, y)
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(x, OneAway)
                Direction = Direction + 1
              End If
            Else
              'We went off the bottom.  Try new direction.
              Direction = Direction + 1
            End If
          Case LEFTY
            OneAway = x - 1
            If OneAway > 0 Then
              'Check and see what's up.
              If ClearDirection(OneAway, y) = True Then
                'Map is clear in that direction.
                x = x - 1
                Call WriteToTent(x, y)
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(OneAway, y)
                Direction = Direction + 1
              End If
            Else
              'We went off the left.  Try new direction.
              Direction = Direction + 1
            End If
          Case RIGHTY
            OneAway = x + 1
            If OneAway <= XDimension Then
              'Check and see what's up.
              If ClearDirection(OneAway, y) = True Then
                'Map is clear in that direction.
                x = x + 1
                Call WriteToTent(x, y)
                Done = True
              Else
                'Blocked.  Try new direction.
                Call GetAdjacentColor(OneAway, y)
                Direction = Direction + 1
              End If
            Else
              'We went off the right.  Try new direction.
              Direction = Direction + 1
            End If
        End Select
        Loop
        If Direction = 5 Then
          'This block is boxed in.  Try the next block in the lake.
          Block = Block + 1
          If Block > TentNum Then
            'Minimum lake can't fit here -- need to fill it in.
            For i = 1 To TentNum
              Map(Tent(1, i), Tent(2, i)) = Country
            Next i
            TentNum = MinLakeSize   'Fudging out of loop.
            FilledIn = True
          End If
        End If
        If Done = True Then
          Block = 1
        End If
      Loop
      If FilledIn = False Then
        'We've reached the minimum lake size.  Now we need to identify this
        'entire body of water with the next lake code.  'Note that we overwrite
        'our coastline here -- we don't need it anymore.
        
        'Part one of this procedure is to fill in the lake part we've
        'found already.
        For i = 1 To TentNum
          Map(Tent(1, i), Tent(2, i)) = LakeCode
        Next i
        
        'Part two of this procedure is to keep filling in the lake until
        'there is no more to fill in.  We do this with quick multiple passes.
        Call LabelBodyOfWater(LakeCode)
        LakeCode = LakeCode + 1
        
      End If
    End If
  Next k
Next j
End Function

Private Sub LabelBodyOfWater(LakeCode As Long)

Dim ii As Long
Dim jj As Long
Dim xx As Long
Dim yy As Long

StillKicking    'Put a period in the commentary line for each lake.

Done = False
Do While Done = False
  Done = True  'If nothing gets written this pass, this will stay.
  For ii = 1 To XDimension
    For jj = 1 To YDimension
      If Map(ii, jj) = LakeCode Then
        'Check for all adjacent squares and fill them in if they are
        'water or coastline.  Note that we do diagonals here!
        xx = ii
        yy = jj
        'Check up.
        If yy - 1 > 0 Then
          If Map(xx, yy - 1) = 0 Or Map(xx, yy - 1) = TILEVAL_COASTLINE Then
            Map(xx, yy - 1) = LakeCode
            Done = False
          End If
          'Check upper left.
          If xx - 1 > 0 Then
            If Map(xx - 1, yy - 1) = 0 Or Map(xx - 1, yy - 1) = TILEVAL_COASTLINE Then
              Map(xx - 1, yy - 1) = LakeCode
              Done = False
            End If
          End If
        End If
        'Check down.
        If yy + 1 <= YDimension Then
          If Map(xx, yy + 1) = 0 Or Map(xx, yy + 1) = TILEVAL_COASTLINE Then
            Map(xx, yy + 1) = LakeCode
            Done = False
          End If
          'Check lower right.
          If xx + 1 <= XDimension Then
            If Map(xx + 1, yy + 1) = 0 Or Map(xx + 1, yy + 1) = TILEVAL_COASTLINE Then
              Map(xx + 1, yy + 1) = LakeCode
              Done = False
            End If
          End If
        End If
        'Check left.
        If xx - 1 > 0 Then
          If Map(xx - 1, yy) = 0 Or Map(xx - 1, yy) = TILEVAL_COASTLINE Then
            Map(xx - 1, yy) = LakeCode
            Done = False
          End If
          'Check lower left.
          If yy + 1 <= YDimension Then
            If Map(xx - 1, yy + 1) = 0 Or Map(xx - 1, yy + 1) = TILEVAL_COASTLINE Then
              Map(xx - 1, yy + 1) = LakeCode
              Done = False
            End If
          End If
        End If
        'Check right.
        If xx + 1 <= XDimension Then
          If Map(xx + 1, yy) = 0 Or Map(xx + 1, yy) = TILEVAL_COASTLINE Then
            Map(xx + 1, yy) = LakeCode
            Done = False
          End If
          'Check upper right.
          If yy - 1 > 0 Then
            If Map(xx + 1, yy - 1) = 0 Or Map(xx + 1, yy - 1) = TILEVAL_COASTLINE Then
              Map(xx + 1, yy - 1) = LakeCode
              Done = False
            End If
          End If
        End If
      End If
    Next jj
  Next ii
Loop

End Sub

Private Function FindNeighbors() As Boolean

Dim i As Long
Dim j As Long
Dim k As Long
Dim x As Long
Dim y As Long

'This routine will populate the Neighbors array.
For i = 1 To NumCountries
  For j = 1 To MAXIMUM_NEIGHBORS
    Neighbor(i, j) = 0
  Next j
Next i

For i = 1 To XDimension
  StillKicking  'Update text.
  For j = 1 To YDimension
    x = i
    y = j
    'See what's up.
    If y - 1 > 0 Then FindNeighbors = AddNeighbor(Map(x, y), Map(x, y - 1))
    'See what's down.
    If y + 1 <= YDimension Then FindNeighbors = AddNeighbor(Map(x, y), Map(x, y + 1))
    'See what's left.
    If x - 1 > 0 Then FindNeighbors = AddNeighbor(Map(x, y), Map(x - 1, y))
    'See what's right.
    If x + 1 <= XDimension Then FindNeighbors = AddNeighbor(Map(x, y), Map(x + 1, y))
    If FindNeighbors = False Then
      'We failed our spec -- too many neighbors.  Let's flag it up the chain.
      Exit Function
    End If
  Next j
Next i

'Now we sort the Neighbor array so that all water is after land.
'This is important when we calculate attack and defense strengths.
'This is just a simple bubble sort.  Not very efficient.    :)
For i = 1 To NumCountries
  Done = False
  Do While Done = False
    Done = True
    For j = (MAXIMUM_NEIGHBORS - 1) To 1 Step -1
      'See if we need to swap two consecutive neighbors, keeping zeros at the end!
      If (Neighbor(i, j + 1) < Neighbor(i, j)) And (Neighbor(i, j + 1) <> 0) Then
        'We swap these two.
        k = Neighbor(i, j + 1)
        Neighbor(i, j + 1) = Neighbor(i, j)
        Neighbor(i, j) = k
        'Mark that we need to go another round.
        Done = False
      End If
    Next j
  Loop
Next i

FindNeighbors = True  'Success bulding the Neighbors array.

End Function

Private Function AddNeighbor(ToCountry As Long, NewNeighbor As Long) As Boolean

'This sub works with FindNeighbors to populate the Neighbors array.

Dim k As Long

AddNeighbor = True  'Success adding this neighbor.

If ToCountry >= TILEVAL_COASTLINE Then Exit Function 'This is water!
If ToCountry = NewNeighbor Then Exit Function 'Can't be our own neighbor!

Dim LastOne As Integer

LastOne = -1

For k = 1 To MAXIMUM_NEIGHBORS
  If Neighbor(ToCountry, k) = NewNeighbor Then Exit Function 'Already there.
  If LastOne = -1 And Neighbor(ToCountry, k) = 0 Then LastOne = k
Next k

If LastOne = -1 Then
  'We have overstepped our limit on the number of allowable neighbors
  'per country.  Let's flag this up the chain so we can rebuild our map.
  AddNeighbor = False
  Exit Function
End If

Neighbor(ToCountry, LastOne) = NewNeighbor

End Function

Private Sub FindSpots()

'This sub will populate the DisplaySpot array.

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim x As Long
Dim y As Long

Dim CurrentCountry As Long
Dim LandCount As Integer

'First, we zero it all out.
For i = 1 To NumCountries
  DisplaySpot(i, 1) = 0
  DisplaySpot(i, 2) = 0
  DisplaySpot(i, 3) = 0
Next i

'We make one pass across the map.  We find the grid location for each
'country that has the most land around it belonging to the same country.
'we check up to 2 squares max in each direction.  Therefore, this number
'can range from 2 to 24.
For i = 1 To Xsize
  StillKicking   'Put a period in the commentary line.
  For j = 1 To Ysize
    CurrentCountry = Map(i, j)
    If CurrentCountry > 0 And CurrentCountry < TILEVAL_COASTLINE Then
      'We really do have a country and not water.
      'Lets check 2 squares in each direction for like land.
      LandCount = 0
      For k = -2 To 2 Step 1
        For l = -2 To 2 Step 1
          x = i + k
          y = j + l
          If x > 0 And x <= Xsize And y > 0 And y <= Ysize Then
            'We aren't off the edge of the map.
            If Map(x, y) = CurrentCountry Then
              'We found a matching piece.  Let's add the correct score
              'according to the following diagram:
              '1 1 1 1 1
              '1 3 3 3 1
              '1 9 X 9 1
              '1 3 3 3 1
              '1 1 1 1 1
              'This way, a square with both left *and* right free has
              'a much better chance of making it.  With both of these free,
              'the number is *much* more readable and there won't be
              'as many conflicts with the borderlines.
              'Note that having one of the horizontally adjacent
              'squares is just as important as having all three to the
              'top or bottom!
              If l = 0 And (k = -1 Or k = 1) Then
                'The big score.
                LandCount = LandCount + 9
              ElseIf (l = -1 Or l = 1) And (k > -2 And k < 2) Then
                'The middle score.
                LandCount = LandCount + 3
              Else
                'Just one point.
                LandCount = LandCount + 1
              End If
            End If
          End If
        Next l
      Next k
      'Now we see if we beat our last high score for this country.
      If LandCount > DisplaySpot(CurrentCountry, 3) Then
        'We do have more than our previous best.  Record this one.
        DisplaySpot(CurrentCountry, 1) = i
        DisplaySpot(CurrentCountry, 2) = j
        DisplaySpot(CurrentCountry, 3) = LandCount
      ElseIf LandCount = DisplaySpot(CurrentCountry, 3) Then
        'We tied our last high score.  Let's add some randomness
        'so that we don't always have the best spot to the upper
        'left of the country.
        If Rnd(1) < 0.05 Then
          '5% chance of taking the new one.
          DisplaySpot(CurrentCountry, 1) = i
          DisplaySpot(CurrentCountry, 2) = j
          DisplaySpot(CurrentCountry, 3) = LandCount
        End If
      End If
    End If
  Next j
Next i

End Sub

Private Sub BuildSeedMatrix()

Dim i As Long
Dim j As Long
Dim marker() As Boolean

ReDim marker(NumCountries)

'Initializing seed matrix.
For i = 1 To NumCountries
  'Using the troops array to flag which countries we've hit.
  marker(i) = 0
Next i
For i = 1 To XDimension
  For j = 1 To YDimension
    If Map(i, j) > 0 And Map(i, j) < TILEVAL_COASTLINE Then
      'if this square is on coastline, and it's the FIRST piece of coastline
      'then we need to force a wooden pier on it.  If not, then go normal.
      If (marker(Map(i, j)) = 0) And (IsOnCoastLine(i, j) = True) Then
        Seed(i, j) = 0
        marker(Map(i, j)) = 1
      Else
        Seed(i, j) = Rnd(1)
      End If
    End If
  Next j
Next i

End Sub

Private Sub DrawNumber(Source As Long, dest As Long, Xcoord As Long, Ycoord As Long, ThreeDigitNumber As Long)

Dim HDigit As Integer
Dim TDigit As Integer
Dim Digit As Integer
Dim Xadj As Long
Dim Yadj As Long
Dim TempNum As Long
Dim TextColor As Boolean

'Get the color of this text.
TempNum = MyMap.Grid(CInt(Xcoord), CInt(Ycoord))   'Get the country ID.
If TempNum > TILEVAL_COASTLINE Then Exit Sub  'Shouldn't get this!
Digit = MyMap.Owner(TempNum)                       'Get the owner.
TextColor = True  'Default to white text.
'Now see if we should have black instead.
If Digit > 0 Then
  If PlayerTextColor(Player(Digit)) = vbBlack Then TextColor = False
Else
  If PlayerTextColor(CFGUnoccupiedColor) = vbBlack Then TextColor = False
End If
'Always black if flashing!
If FlashArray(TempNum, 1) = 1 Then TextColor = False

TempNum = ThreeDigitNumber

'Get the three digits.
HDigit = Int(TempNum / 100)
TempNum = TempNum - (HDigit * 100)
TDigit = Int(TempNum / 10)
TempNum = TempNum - (TDigit * 10)
Digit = TempNum

Yadj = (Ycoord - 1) * 8

'Hundreds digit.  Offset 5 pixels left if it exists.
If (HDigit > 0) Then
  Xadj = ((Xcoord - 1) * GFX_GRID) - 5
  Call PlaceDigit(Source, dest, HDigit, Xadj, Yadj, TextColor)
End If

'Tens digit.
If (HDigit > 0) Or (TDigit > 0) Then
  If HDigit = 0 Then
    'If less than 100, we offset by 3 pixels left.
    Xadj = ((Xcoord - 1) * GFX_GRID) - 3
  Else
    'No offset.
    Xadj = ((Xcoord - 1) * GFX_GRID)
  End If
  Call PlaceDigit(Source, dest, TDigit, Xadj, Yadj, TextColor)
End If

'Ones digit.
If (HDigit > 0) Or (TDigit > 0) Or (Digit > 0) Then
  If HDigit = 0 And TDigit = 0 Then
    'If less than 10, we center it.
    Xadj = ((Xcoord - 1) * GFX_GRID)
  ElseIf HDigit = 0 And TDigit > 0 Then
    'If less than 100, we offset by 2 pixels right.
    Xadj = ((Xcoord - 1) * GFX_GRID) + 2
  Else
    'Offset 5 pixels right.
    Xadj = ((Xcoord - 1) * GFX_GRID) + 5
  End If
  Call PlaceDigit(Source, dest, Digit, Xadj, Yadj, TextColor)
End If

End Sub

Private Sub PlaceDigit(Source As Long, dest As Long, Digit As Integer, Xadj As Long, Yadj As Long, BlackAndWhite As Boolean)

'Draw a single number on the screen.
u% = BitBlt(dest, Xadj, Yadj, GFX_GRID_HALF, GFX_GRID, Source, (GFX_DIGIT_X_BASE + Digit) * GFX_GRID, GFX_WHITE_Y_COORD, SRCAND)
If BlackAndWhite = False Then
  'Black.
  u% = BitBlt(dest, Xadj, Yadj, GFX_GRID_HALF, GFX_GRID, Source, GFX_BLACK_X_COORD, GFX_HQ_Y_COORD, SRCINVERT)
Else
  'White.
  u% = BitBlt(dest, Xadj, Yadj, GFX_GRID_HALF, GFX_GRID, Source, (GFX_DIGIT_X_BASE + Digit) * GFX_GRID, GFX_HQ_Y_COORD, SRCINVERT)
End If

End Sub

Public Sub SetupTitleScreen()

Dim i As Long
Dim j As Long
Dim Xcoo As Integer
Dim Ycoo As Integer
Dim MapItem As Integer

'Everything is water...
For i = 1 To XDimension
  For j = 1 To YDimension
    Map(i, j) = 1337
    Seed(i, j) = Rnd(1)
  Next j
Next i

'Get which title we're supposed to be showing.
SetupTitleScreenLand

'Except where there's an "X" in our template.
For j = 1 To CurrentTitleHeight
  For i = 1 To CurrentTitleWidth
    Select Case Mid(TitleLand(j), i, 1)
        Case "A"
            'This square of the map has land on it!  Let's paint it in.
            MapItem = 2
        Case "B"
            'This square of the map has land on it!  Let's paint it in.
            MapItem = 3
        Case "C"
            'This square of the map has land on it!  Let's paint it in.
            MapItem = 4
        Case "X"
            'This square of the map has land on it!  Let's paint it in.
            MapItem = 1
        Case Else
            'Make it water by default.
            MapItem = 1337
    End Select
    
    'Find the map coordinate where this should go.  Center it and bounds check.
    Xcoo = Int(XDimension / 2) - Int(CurrentTitleWidth / 2) + i
    Ycoo = Int(YDimension / 2) - Int(CurrentTitleHeight / 2) + j
    If Xcoo < 1 Then Xcoo = 1
    If Xcoo > XDimension Then Xcoo = XDimension
    If Ycoo < 1 Then Ycoo = 1
    If Ycoo > YDimension Then Ycoo = YDimension
    
    'Finally, add this square to our map grid.
    Map(Xcoo, Ycoo) = MapItem
    
  Next i
Next j

ReDim Special(4)
ReDim CountryClr(4)
ReDim FlashArray(4, 4)

CountryClr(1) = CFGUnoccupiedColor
CountryClr(2) = 1
CountryClr(3) = 3
CountryClr(4) = 5
Special(1) = ((Rnd(1) < 0.2) * 1) + ((Rnd(1) < 0.5) * 2)
Special(2) = ((Rnd(1) < 0.5) * 1) + ((Rnd(1) < 0.2) * 2)
Special(3) = ((Rnd(1) < 0.5) * 1) + ((Rnd(1) < 0.2) * 2)
Special(4) = ((Rnd(1) < 0.5) * 1) + ((Rnd(1) < 0.2) * 2)

End Sub

Public Sub RedimensionStuff(xxx As Long, yyy As Long)

On Error GoTo ErrStuff

'Redimension our arrays to save some memory.
NumCountries = xxx
LakeCode = yyy

ReDim Neighbor(NumCountries, MAXIMUM_NEIGHBORS)
ReDim CountryClr(NumCountries)
ReDim Owners(NumCountries)
ReDim DisplaySpot(NumCountries, 3)
ReDim Troops(NumCountries)
ReDim Special(NumCountries)
ReDim CName(NumCountries)
ReDim WName(LakeCode - 1000)
ReDim FlashArray(NumCountries, 4)

ErrStuff:

End Sub

Public Sub FinishUpLoad()

Dim x As Boolean

On Error GoTo ErrStuff

'FindNeighbors *has* to be true, or we never would have built the map
'in the first place.
x = FindNeighbors

Call FindSpots
Call BuildSeedMatrix

'We should have loaded the stamp from file.  For a net game, this doesn't matter.
MyMap.MapStamp = LastMapStamp

ErrStuff:

End Sub

Public Function UpdateFlashing(Cntry As Long) As Boolean

'Update the flash array for the passed country.  If we triggered
'a flash, flag the caller to redraw the screen.

UpdateFlashing = False
If FlashArray(Cntry, 2) > 0 Then
  'This country is flashing now.  Decrement the counter.
  FlashArray(Cntry, 2) = FlashArray(Cntry, 2) - 1
  If FlashArray(Cntry, 2) = 0 Then
    'Reverse our status...
    FlashArray(Cntry, 1) = 1 - FlashArray(Cntry, 1)
    'Set the flag...
    UpdateFlashing = True
    'One flash down...
    FlashArray(Cntry, 4) = FlashArray(Cntry, 4) - 1
    If FlashArray(Cntry, 4) > 0 Then
      'Still more to go.  Reset for next.
      FlashArray(Cntry, 2) = FlashArray(Cntry, 3)
    End If
  End If
End If

End Function

Private Function CorrectColor(Cntry As Long) As Long

'This sub either returns the color of the specified country,
'or white if it's flashing!
If FlashArray(Cntry, 1) = 0 Then
  'We'll use our normal color -- we're not flashing.
  CorrectColor = CountryClr(Cntry)
Else
  'This country is flashing!  It's white by default.
  CorrectColor = 12
End If

End Function

Private Sub StillKicking()

'This sub puts a period on the end of the commentary to show that the
'map maker is still working..........

If Len(Land.Commentary.Caption) > 80 Then
  Land.Commentary.Caption = "Creating map..."
  Land.Oopsie.Caption = ""
  Call Land.Oopsie.Refresh
Else
  Land.Commentary.Caption = Land.Commentary.Caption & "."
End If
Call Land.Commentary.Refresh

End Sub

Public Sub SetDimensions(x As Integer, y As Integer)

XDimension = x
YDimension = y

End Sub
