Attribute VB_Name = "LoadAndSave"
Option Explicit
Option Base 1

Public Function LoadMap(OptionsOnly As Boolean) As Integer

Dim MapFile As String

On Error GoTo PressCancel

LoadMap = 0

'Set up the common dialog control.
Land.CD1.DialogTitle = "Load Map"
Land.CD1.FileName = ""
Land.CD1.Filter = "Map Files (.map)|*.map|All Files|*.*"
Land.CD1.FilterIndex = 1
Land.CD1.InitDir = MAPpath
Land.CD1.Flags = &H8 Or &H4
Land.CD1.CancelError = True
'This displays it.
Land.CD1.ShowOpen

'If we got here, then the user didn't cancel the dialog.
MapFile = Land.CD1.FileName

'If the file the user chose doesn't exist, we exit.
If Dir(MapFile) = "" Then
  MsgBox "The specified map could not be found.  Please check the filename and try again.", vbOKOnly Or vbCritical, "File Not Found"
  Exit Function
End If

LoadMap = ReadMapDataKitchenSink(MapFile, OptionsOnly)

Exit Function

PressCancel:

End Function

Public Sub SaveMap()

Dim MapFile As String

On Error GoTo PressedCancel

'Set up common dialog control.
Land.CD1.DialogTitle = "Save Current Map"
If MyMap.MapName <> vbNullString Then
  Land.CD1.FileName = RawMapName(MyMap.MapName) & ".map"
Else
  Land.CD1.FileName = "Fracas.map"
End If
Land.CD1.Filter = "Map Files (.map)|*.map|All Files|*.*"
Land.CD1.FilterIndex = 1
Land.CD1.InitDir = MAPpath
Land.CD1.Flags = &H2 Or &H4 Or &H8
Land.CD1.CancelError = True
'This displays it.
Land.CD1.ShowSave

'If we got here, then the user didn't cancel the dialog.
MapFile = Land.CD1.FileName

'First, put all of our INI settings in the save file.
MakeINI MapFile

'Assign our file name to the map.  In case the file is moved, this will
'be updated on the next save for whatever reason.
MyMap.MapName = MapFile

'Now write physical map data to the file.
WriteMapData MapFile

Exit Sub

PressedCancel:

End Sub

Public Function LoadGame() As Integer

Dim GameFile As String
Dim GameFileNumber As Long
Dim MyMapPath As String
Dim MyMapStamp As String
Dim CurrentLine As String

On Error GoTo PressCancel

LoadGame = 0

'Set up the common dialog control.
Land.CD1.DialogTitle = "Load Saved Game"
Land.CD1.FileName = ""
Land.CD1.Filter = "Save Files (.sav)|*.sav|All Files|*.*"
Land.CD1.FilterIndex = 1
Land.CD1.InitDir = MAPpath
Land.CD1.Flags = &H8 Or &H4
Land.CD1.CancelError = True
'This displays it.
Land.CD1.ShowOpen

'If we got here, then the user didn't cancel the dialog.
GameFile = Land.CD1.FileName

'If the file the user chose doesn't exist, we exit.
If Dir(GameFile) = "" Then
  MsgBox "The specified game could not be found.  Please check the filename and try again.", vbOKOnly Or vbCritical, "File Not Found"
  Exit Function
End If

'Grab the map path and timestamp from the saved game.
GameFileNumber = FreeFile
Open GameFile For Input As GameFileNumber

'Get to the map info.
MyMapPath = vbNullString
Do While Not EOF(GameFileNumber)
  Line Input #GameFileNumber, CurrentLine
  If CurrentLine = "{Map Path}" Then
    'Grab the path to the map we apply to.
    Line Input #GameFileNumber, CurrentLine
    MyMapPath = CurrentLine
    Line Input #GameFileNumber, CurrentLine
    MyMapStamp = Arg(StringAfterEqual(CurrentLine), 1)
    Exit Do
  End If
Loop

Close #GameFileNumber

'We found a path, now verify that it is valid.
If MyMapPath <> "" Then
  If Dir(MyMapPath) <> "" Then
    'We have a file there, so let's grab all information in it.
    LoadGame = ReadMapDataKitchenSink(MyMapPath, False)
    'If the stamp in the map doesn't match the stamp in the saved game, return 0.
    If MyMapStamp <> MyMap.MapStamp Then
      LoadGame = 0
      MsgBox "This saved game does not apply to this map.", vbOKOnly Or vbExclamation, "Timestamp Mismatch"
      Exit Function
    End If
    'Now that our map structure is in place, populate all game data from file.
    If LoadGame > 0 Then LoadGame = ReadGameData(GameFile)
  Else
    MsgBox "Could not locate the map to which this saved game applies.", vbOKOnly Or vbExclamation, "Map Not Found"
    Exit Function
  End If
Else
  MsgBox "Could not locate the map to which this saved game applies.", vbOKOnly Or vbExclamation, "Map Not Found"
  Exit Function
End If

Exit Function

PressCancel:

End Function

Public Function ReadMapDataKitchenSink(MapFile As String, OptionsOnly As Boolean) As Integer

'Called by load map and load game routines to grab everything from the map file.
'At this point we've verified that MapFile is good.
ReadMapDataKitchenSink = 0
  
'Get all the menu settings in the map...
ReadINI MapFile

'Update our menus to reflect the saved config settings.
Call ClearMenuChecks
Call UpdateMenus

'Now that we have a new resolution, see if it's valid.
If CFGResolution > MaxAllowedResolution Then
  'Can't display a map at that resolution.  Exit gracefully.
  MsgBox "Saved map size is larger than the desktop and cannot be used.", vbOKOnly Or vbCritical, "Map Too Large"
  Exit Function
Else
  'Resolution is valid, set up map.
  Land.ChangeRes
End If

If OptionsOnly = False Then
  'Get the physical map information from the same file.
  Call ReadMapData(MapFile)
  'Some map parameters can be calculated instead of saved, so do these now.
  Call MyMap.FinishUpLoad
  'Recalculate ShipPct.  Fudge.
  Call Land.CollectMenuSettings(True)
End If

ReadMapDataKitchenSink = 1

End Function

Public Sub SaveGame(Turn As Integer, Phase As Integer)

'This sub saves the current game.  Note that we keep track of which map this
'saved game applies to.

Dim GameFile As String

On Error GoTo PressedCancel

'Set up common dialog control.
Land.CD1.DialogTitle = "Save Current Game"
Land.CD1.FileName = RawMapName(MyMap.MapName) & ".sav"
Land.CD1.Filter = "Save Files (.sav)|*.sav|All Files|*.*"
Land.CD1.FilterIndex = 1
Land.CD1.InitDir = MAPpath
Land.CD1.Flags = &H2 Or &H4 Or &H8
Land.CD1.CancelError = True
'This displays it.
Land.CD1.ShowSave

'If we got here, then the user didn't cancel the dialog.
GameFile = Land.CD1.FileName

'We will update the map file as well to take into account things
'that may have changed like country names, etc.
QuickMapUpdate

'Now write game data to the save file.
WriteGameData GameFile, Turn, Phase

Exit Sub

PressedCancel:

End Sub

Public Sub QuickMapUpdate()

Dim MapFile As String

'Quickly rewrite our map file.
MapFile = MyMap.MapName
MakeINI MapFile
WriteMapData MapFile

End Sub

Private Sub ReadMapData(MapFile As String)

'This sub reads MAP data from the passed file path.  MAP data is the physical
'structure of the map -- its dimensions, land/water squares, and names.

On Error GoTo FileError

Dim CurrentLine As String
Dim MapFileNumber As Long
Dim TempNum1 As Long
Dim TempNum2 As Long
Dim si As Integer
Dim sj As Integer
Dim sk As Long
Dim SepPoint As Long

'Menu settings are done, now grab map info.
MapFileNumber = FreeFile
Open MapFile For Input As MapFileNumber

'Get to the map info.
Do While Not EOF(MapFileNumber)
  Line Input #MapFileNumber, CurrentLine
  If CurrentLine = "{Menu and Map Data}" Then

    'Grab the governing numbers.
    Line Input #MapFileNumber, CurrentLine
    'Note that we don't care what path we get in the above line.
    Line Input #MapFileNumber, CurrentLine
    TempNum1 = Val(CurrentLine)
    Line Input #MapFileNumber, CurrentLine
    TempNum2 = Val(CurrentLine)
    'Redim arrays.
    Call MyMap.RedimensionStuff(TempNum1, TempNum2)
    
    'Input map contents.
    Line Input #MapFileNumber, CurrentLine   'Text spacer.
    For sj = 1 To MyMap.Ysize
      CurrentLine = ""
      Line Input #MapFileNumber, CurrentLine
      If CurrentLine <> "" Then
        For si = 1 To MyMap.Xsize
          'Grab each coordinate separately...
          SepPoint = InStr(CurrentLine, ".")
          MyMap.Grid(si, sj) = Val(Left(CurrentLine, SepPoint - 1))
          CurrentLine = Right(CurrentLine, Len(CurrentLine) - SepPoint)
        Next si
      End If
    Next sj
    
    'Input country names.
    Line Input #MapFileNumber, CurrentLine   'Text spacer.
    For sk = 1 To TempNum1
      CurrentLine = ""
      Line Input #MapFileNumber, CurrentLine
      MyMap.CountryName(sk) = CurrentLine
    Next sk
    
    'Input water mass names.
    Line Input #MapFileNumber, CurrentLine   'Text spacer.
    If TempNum2 > 1001 Then
      For sk = 1 To (TempNum2 - 1000)
        CurrentLine = ""
        Line Input #MapFileNumber, CurrentLine
        MyMap.WaterName(sk) = CurrentLine
      Next sk
    End If
    
    'Input hi scores.
    Line Input #MapFileNumber, CurrentLine   'Text spacer.
    For sk = 1 To NUM_HI_SCORES
      CurrentLine = ""
      Line Input #MapFileNumber, CurrentLine
      HiScoreName(sk) = Arg(StringAfterEqual(CurrentLine), 1)
      HiScore(sk) = Val(Arg(StringAfterEqual(CurrentLine), 2))
      HiScoreColor(sk) = Val(Arg(StringAfterEqual(CurrentLine), 3))
    Next sk
    
  End If

Loop

Close #MapFileNumber

'If we got here, then the load was a success!  Set our new file name.
MyMap.MapName = MapFile

Exit Sub

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "ReadMapData - " & Err.Number & " - " & Err.Description

End Sub

Private Sub WriteMapData(MapFile As String)

'This sub writes all physical map data to the passed file.
'Physical data includes country dimensions and placement, names, and high scores.

On Error GoTo FileError

Dim MapFileNumber As Long
Dim CurrentLine As String
Dim si As Integer
Dim sj As Integer
Dim sk As Long

MapFileNumber = FreeFile
Open MapFile For Append As MapFileNumber

Print #MapFileNumber, ""

'Write governing numbers.
Print #MapFileNumber, "{Menu and Map Data}"
Print #MapFileNumber, MyMap.MapName
Print #MapFileNumber, MyMap.NumberOfCountries
Print #MapFileNumber, MyMap.LakeCode

'Write the entire map contents.
Print #MapFileNumber, "{Map}"
For sj = 1 To MyMap.Ysize
  CurrentLine = ""
  For si = 1 To MyMap.Xsize
    CurrentLine = CurrentLine & Trim(Str(MyMap.Grid(si, sj))) & "."
  Next si
  Print #MapFileNumber, CurrentLine
Next sj

'Write the country names.
Print #MapFileNumber, "{Country Names}"
For sk = 1 To MyMap.NumberOfCountries
  Print #MapFileNumber, MyMap.CountryName(sk)
Next sk

'Write the water mass names.
Print #MapFileNumber, "{Water Names}"
If MyMap.LakeCode > 1001 Then
  For sk = 1 To MyMap.LakeCode - 1000
    Print #MapFileNumber, MyMap.WaterName(sk)
  Next sk
End If

'Write the high scores.
Print #MapFileNumber, "{Hi Scores}"
For sk = 1 To NUM_HI_SCORES
  Print #MapFileNumber, "HI" & Trim(Str(sk)) & "=" & Trim(HiScoreName(sk)) & _
            "," & Trim(Str(HiScore(sk))) & "," & Trim(Str(HiScoreColor(sk)))
Next sk

Close #MapFileNumber

Exit Sub

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "WriteMapData - " & Err.Number & " - " & Err.Description

End Sub

Private Function ReadGameData(GameFile As String) As Integer

'This sub reads MAP data from the passed file path.  MAP data is the physical
'structure of the map -- its dimensions, land/water squares, and names.

On Error GoTo FileError

ReadGameData = 0

Dim CurrentLine As String
Dim TempStr As String
Dim TempInt As Integer
Dim TempLong As Long
Dim GameFileNumber As Long
Dim i As Integer
Dim j As Integer

'Grab map info.
GameFileNumber = FreeFile
Open GameFile For Input As GameFileNumber

'At this point, we've read our map data.  Now finish with the saved game.
Do While Not EOF(GameFileNumber)
  Line Input #GameFileNumber, CurrentLine
  If Left(CurrentLine, 1) <> ";" Then   'Ignore comments.
    TempStr = StringBeforeEqual(CurrentLine)
    If TempStr <> "" Then   'Ignore lines without an assignment.
      TempInt = Val(StringAfterEqual(CurrentLine))
      Select Case TempStr
        Case "Turn": Land.SetTurn TempInt
        Case "Phase":  Land.SetPhase TempInt
        Case "TurnCounter": TurnCounter = TempInt
        'Player data.
        Case "Player1", "Player2", "Player3", "Player4", "Player5", "Player6":
          TempInt = Val(Right(TempStr, 1))  'Take the number off the end...
          NumOccupied(TempInt) = Val(Arg(StringAfterEqual(CurrentLine), 1))
          NumTroops(TempInt) = Val(Arg(StringAfterEqual(CurrentLine), 2))
      End Select
      'Country data.
      If Left(TempStr, 1) = "C" Then
        TempLong = Val(Right(TempStr, Len(TempStr) - 1))  'Strip off the C.
          MyMap.Owner(TempLong) = Val(Arg(StringAfterEqual(CurrentLine), 1))
          MyMap.TroopCount(TempLong) = Val(Arg(StringAfterEqual(CurrentLine), 2))
          MyMap.CountryType(TempLong) = Val(Arg(StringAfterEqual(CurrentLine), 3))
          MyMap.CountryColor(TempLong) = Val(Arg(StringAfterEqual(CurrentLine), 4))
      End If
      'Stats.
      If Left(TempStr, 4) = "STAT" Then
        i = Val(Right(TempStr, 1))  'Get the player number...
        For j = 1 To MAX_PLAYERS
          STATattacked(i, j) = Val(Arg(StringAfterEqual(CurrentLine), ((j - 1) * 4) + 1))
          STATovertaken(i, j) = Val(Arg(StringAfterEqual(CurrentLine), ((j - 1) * 4) + 2))
          STATkilled(i, j) = Val(Arg(StringAfterEqual(CurrentLine), ((j - 1) * 4) + 3))
          STATdefeated(i, j) = Val(Arg(StringAfterEqual(CurrentLine), ((j - 1) * 4) + 4))
        Next j
      End If
    End If
  End If
Loop

Close #GameFileNumber

ReadGameData = 1

Exit Function

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "ReadGameData - " & Err.Number & " - " & Err.Description

End Function

Private Sub WriteGameData(GameFile As String, Turn As Integer, Phase As Integer)

'This sub writes relevant game data to the passed file.

On Error GoTo FileError

Dim GameFileNumber As Long
Dim i As Long
Dim si As Integer
Dim sj As Integer
Dim TempStr As String

If Dir(GameFile) <> "" Then Kill GameFile

GameFileNumber = FreeFile
Open GameFile For Output As GameFileNumber

Print #GameFileNumber, ";" & VersionStr     'Version info.
Print #GameFileNumber, ";" & WebStr         'Shameless self-promotion.
Print #GameFileNumber, ";" & EmailStr       'More shameless self-promotion.

'Write the path of this map first.  This is the map we'll apply these
'game parameters to when we load.
Print #GameFileNumber, ""
Print #GameFileNumber, "{Map Path}"
Print #GameFileNumber, MyMap.MapName
Print #GameFileNumber, "MatchToStamp=" & MyMap.MapStamp
Print #GameFileNumber, ""

'Write Game parameters.
Print #GameFileNumber, "{Game Data}"
Print #GameFileNumber, "Turn=" & Trim(Str(Turn))
Print #GameFileNumber, "Phase=" & Trim(Str(Phase))
Print #GameFileNumber, "TurnCounter=" & Trim(Str(TurnCounter))

'Write Country data.
Print #GameFileNumber, ""
Print #GameFileNumber, "{Country Data}"
For i = 1 To MyMap.NumberOfCountries
  Print #GameFileNumber, "C" & Trim(Str(i)) & "=" & _
          Trim(Str(MyMap.Owner(i))) & "," & _
          Trim(Str(MyMap.TroopCount(i))) & "," & _
          Trim(Str(MyMap.CountryType(i))) & "," & _
          Trim(Str(MyMap.CountryColor(i)))
Next i

'Write Player data.
Print #GameFileNumber, ""
Print #GameFileNumber, "{Player Data}"
For i = 1 To 6
  Print #GameFileNumber, "Player" & Trim(Str(i)) & "=" & _
          Trim(Str(NumOccupied(i))) & "," & _
          Trim(Str(NumTroops(i)))
Next i

'Write statistical info.
Print #GameFileNumber, ""
Print #GameFileNumber, "{Statistics}"
For si = 1 To MAX_PLAYERS
  TempStr = vbNullString
  For sj = 1 To MAX_PLAYERS
    TempStr = TempStr & Trim(Str(STATattacked(si, sj))) & "," & _
                        Trim(Str(STATovertaken(si, sj))) & "," & _
                        Trim(Str(STATkilled(si, sj))) & "," & _
                        Trim(Str(STATdefeated(si, sj))) & ","

  Next sj
  Print #GameFileNumber, "STAT" & Trim(Str(si)) & "=" & Left(TempStr, Len(TempStr) - 1)
Next si

Close #GameFileNumber

Exit Sub

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "WriteGameData - " & Err.Number & " - " & Err.Description

End Sub

Public Sub ReadINI(INIFile As String)

'This routine will see if the passed file exists in the root folder.  If it does,
'then we read in all of the config parameters.  If not, we create a default.
'We call this the INI file, even though we also use this routine to read
'menu settings from saved maps and saved games.

Dim INIFileNumber As Integer
Dim TempStr As String
Dim TempInt As Integer
Dim CurrentLine As String
Dim PlNum As Integer

'On Error GoTo FileError

INIFileNumber = FreeFile

'If the file doesn't exist, we exit.
If Dir(INIFile) = "" Then
  'No INI file exists, so create the default.
  If Right(INIFile, 4) = ".ini" Then
    MakeINI INIFile
  End If
End If

'If we got here, then the .ini file should exist, whether we created it or not.
'Now parse it all out.

Open INIFile For Input As INIFileNumber

Do While Not EOF(INIFileNumber)
  Line Input #INIFileNumber, CurrentLine
  If Left(CurrentLine, 1) <> ";" Then   'Ignore comments.
    TempStr = StringBeforeEqual(CurrentLine)
    If TempStr <> "" Then   'Ignore lines without an assignment.
      TempInt = Val(StringAfterEqual(CurrentLine))
      Select Case TempStr
        Case "CountrySize": CFGCountrySize = TempInt
        Case "CountryProportions": CFGProportion = TempInt
        Case "CountryShapes": CFGShape = TempInt
        Case "MinLakeSize": CFGLakeSize = TempInt
        Case "LandPct": CFGLandPct = TempInt
        Case "Islands": CFGIslands = TempInt
        Case "InitTroopPl": CFGInitTroopPl = TempInt
        Case "InitTroopCt": CFGInitTroopCt = TempInt
        Case "BonusTroops": CFGBonusTroops = TempInt
        Case "Ships": CFGShips = TempInt
        Case "Ports": CFGPorts = TempInt
        Case "Conquer": CFGConquer = TempInt
        Case "Events": CFGEvents = TempInt
        Case "1stTurn": CFG1st = TempInt
        Case "HQSelect": CFGHQSelect = TempInt
        Case "Borders": CFGBorders = TempInt
        Case "Sound": CFGSound = TempInt
        Case "AISpeed": CFGAISpeed = TempInt
        Case "AnimSpeed": Land.BallTimer.Interval = TempInt
        Case "Explosions": CFGExplosions = TempInt
        Case "Waves": CFGWaves = TempInt
        Case "UnoccupiedColor": CFGUnoccupiedColor = TempInt
        Case "Flashing": CFGFlashing = TempInt
        Case "Resolution": CFGResolution = TempInt
        Case "Prompt": CFGPrompt = TempInt
        'Player data.
        Case "Player1", "Player2", "Player3", "Player4", "Player5", "Player6":
          PlNum = Val(Right(TempStr, 1))  'Take the number off the end...
          PlayerName(PlNum) = Arg(StringAfterEqual(CurrentLine), 1)
          Player(PlNum) = Val(Arg(StringAfterEqual(CurrentLine), 2))
          PlayerType(PlNum) = Val(Arg(StringAfterEqual(CurrentLine), 3))
          Personality(PlNum) = Val(Arg(StringAfterEqual(CurrentLine), 4))
          Land.Menu1st(PlNum).Caption = PlayerName(PlNum)
        'Time stamp.
        Case "Created":
          LastMapStamp = Arg(StringAfterEqual(CurrentLine), 1)
      End Select
    End If
  End If
Loop

Close INIFileNumber

'Now update our menus with the new parameters.
UpdateMenus

Exit Sub

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "ReadINI - " & Err.Number & " - " & Err.Description

End Sub

Public Sub MakeINI(INIFile As String)

'This sub creates the passed file from scratch and puts menu settings in it.
'It will always use whatever our current config settings are.
'THIS IS ALWAYS THE FIRST THING IN A FRACAS FILE, whether it's a saved map,
'saved game, or the .ini file.

Dim CurrentLine As String
Dim i As Integer
Dim INIFileNumber As Integer
Dim TempType As Integer

On Error GoTo FileError

'If we've never saved this map before, time stamp it.
If Right(INIFile, 4) <> ".ini" Then
  If MyMap.MapStamp = vbNullString Then
    MyMap.MapStamp = Trim(Str(Date) & " " & Str(Time))
  End If
End If

INIFileNumber = FreeFile

If Dir(INIFile) <> "" Then Kill INIFile

Open INIFile For Output As INIFileNumber

Print #INIFileNumber, ";" & VersionStr     'Version info.
Print #INIFileNumber, ";" & WebStr         'Shameless self-promotion.
Print #INIFileNumber, ";" & EmailStr       'More shameless self-promotion.
If Right(INIFile, 4) <> ".ini" Then
  'Don't do this if we're writing the Fracas.ini file.
  Print #INIFileNumber, "Created=" & MyMap.MapStamp
End If

'Now write the contents of the menus.
Print #INIFileNumber,
Print #INIFileNumber, ";Terraform"
Print #INIFileNumber, "CountrySize=" & Trim(Str(CFGCountrySize))
Print #INIFileNumber, "CountryProportions=" & Trim(Str(CFGProportion))
Print #INIFileNumber, "CountryShapes=" & Trim(Str(CFGShape))
Print #INIFileNumber, "MinLakeSize=" & Trim(Str(CFGLakeSize))
Print #INIFileNumber, "LandPct=" & Trim(Str(CFGLandPct))
Print #INIFileNumber, "Islands=" & Trim(Str(CFGIslands))
Print #INIFileNumber,
Print #INIFileNumber, ";Options"
Print #INIFileNumber, "InitTroopPl=" & Trim(Str(CFGInitTroopPl))
Print #INIFileNumber, "InitTroopCt=" & Trim(Str(CFGInitTroopCt))
Print #INIFileNumber, "BonusTroops=" & Trim(Str(CFGBonusTroops))
Print #INIFileNumber, "Ships=" & Trim(Str(CFGShips))
Print #INIFileNumber, "Ports=" & Trim(Str(CFGPorts))
Print #INIFileNumber, "Conquer=" & Trim(Str(CFGConquer))
Print #INIFileNumber, "Events=" & Trim(Str(CFGEvents))
Print #INIFileNumber, "1stTurn=" & Trim(Str(CFG1st))
Print #INIFileNumber, "HQSelect=" & Trim(Str(CFGHQSelect))
Print #INIFileNumber,
Print #INIFileNumber, ";Preferences"
Print #INIFileNumber, "Borders=" & Trim(Str(CFGBorders))
Print #INIFileNumber, "Sound=" & Trim(Str(CFGSound))
Print #INIFileNumber, "AISpeed=" & Trim(Str(CFGAISpeed))
Print #INIFileNumber, "AnimSpeed=" & Trim(Str(Land.BallTimer.Interval))
Print #INIFileNumber, "Explosions=" & Trim(Str(CFGExplosions))
Print #INIFileNumber, "Waves=" & Trim(Str(CFGWaves))
Print #INIFileNumber, "UnoccupiedColor=" & Trim(Str(CFGUnoccupiedColor))
Print #INIFileNumber, "Flashing=" & Trim(Str(CFGFlashing))
Print #INIFileNumber, "Resolution=" & Trim(Str(CFGResolution))
Print #INIFileNumber, "Prompt=" & Trim(Str(CFGPrompt))

'Write Player info.
Print #INIFileNumber,
Print #INIFileNumber, ";Players"
For i = 1 To 6
  'Clients will use the TempPlayerType array instead of PlayerType because PlayerType
  'has to be changed during a network game.
  If MyNetworkRole = NW_CLIENT Then
    TempType = TempPlayerType(i)
  Else
    TempType = PlayerType(i)
  End If
  'Write the line.
  Print #INIFileNumber, "Player" & Trim(Str(i)) & "=" & PlayerName(i) & "," & Player(i) & "," & TempType & "," & Personality(i)
Next i

Close #INIFileNumber

Exit Sub

FileError:

  MsgBox "A file error occurred.  Please Email jmerlo@austin.rr.com with this code:" & vbCrLf & "MakeINI - " & Err.Number & " - " & Err.Description
  
End Sub

Private Function StringBeforeEqual(MyStr As String) As String

'This function returns the first part of a string,
'up to but not including the equals sign.
Dim Pos As Integer

Pos = InStr(MyStr, "=")
If Pos = 0 Then
  StringBeforeEqual = ""
Else
  StringBeforeEqual = Left(MyStr, Pos - 1)
End If

End Function

Private Function StringAfterEqual(MyStr As String) As String

'This function returns the last part of a string,
'everything after the equals sign.

Dim Pos As Integer

Pos = InStr(MyStr, "=")
If Pos = 0 Then
  StringAfterEqual = ""
Else
  StringAfterEqual = Right(MyStr, Len(MyStr) - Pos)
End If

End Function

Public Function Arg(MyStr As String, ArgNum As Integer) As String

'This function returns the specified argument of the passed string.
'Arguments are separated by commas.

Dim i As Integer
Dim Pos As Integer
Dim TempStr As String

TempStr = MyStr
For i = 1 To (ArgNum - 1)
  Pos = InStr(TempStr, ",")
  If Pos = 0 Then
    'Either this is the last argument or there was only one.
    'Either way, there's nothing more to do, so return what we have.
    Arg = TempStr
    Exit Function
  Else
    'We found a comma.  Cut off the first argument up to it.
    TempStr = Right(TempStr, Len(TempStr) - Pos)
  End If
Next i
'Now the string we want is the first argument in the string.
Pos = InStr(TempStr, ",")
If Pos = 0 Then
  'We're at the correct last argument.  Return it.
  Arg = TempStr
Else
  'Only return the string before the comma.
  Arg = Left(TempStr, Pos - 1)
End If

End Function

Public Function ArgAt(MyStr As String, ArgNum As Integer) As String

'This function returns the specified argument of the passed string.
'Arguments are separated by at signs (@).

Dim i As Integer
Dim Pos As Integer
Dim TempStr As String

TempStr = MyStr
For i = 1 To (ArgNum - 1)
  Pos = InStr(TempStr, "@")
  If Pos = 0 Then
    'Either this is the last argument or there was only one.
    'Either way, there's nothing more to do, so return what we have.
    ArgAt = TempStr
    Exit Function
  Else
    'We found a comma.  Cut off the first argument up to it.
    TempStr = Right(TempStr, Len(TempStr) - Pos)
  End If
Next i
'Now the string we want is the first argument in the string.
Pos = InStr(TempStr, "@")
If Pos = 0 Then
  'We're at the correct last argument.  Return it.
  ArgAt = TempStr
Else
  'Only return the string before the comma.
  ArgAt = Left(TempStr, Pos - 1)
End If

End Function

Private Function RawMapName(Mname As String) As String

'This sub extracts JUST the map name from the passed path.  In other words,
'we return whatever is between the last \ and the last . (if it has a .)
Dim Pos1 As Long
Dim Pos2 As Long

Pos1 = InStrRev(Mname, "\") 'String position of the last backslash.
Pos2 = InStrRev(Mname, ".") 'String position of the last period.

If (Pos1 <> 0) And (Pos2 <> 0) And (Pos1 < Pos2) Then
  'There *is* a backslash, and a period, and the period is after the last backslash.
  RawMapName = Mid(Mname, Pos1 + 1, (Pos2 - Pos1 - 1))
ElseIf (Pos1 <> 0) And (Pos2 = 0) Then
  'Backslash, but probably no period.  Return everything after the backslash.
  RawMapName = Right(Mname, Len(Mname) - Pos1)
Else
  'Punt.  Return the whole thing!  I don't know what to do with it.
  RawMapName = Mname
End If

End Function

Public Sub UpdateMenus()

'This subroutine updates the menus with the cfg settings we have.

'First, clear out all menu checks.
ClearMenuChecks

'Now, populate each one with the appropriate check.
Land.MenuBorders(CFGBorders).Checked = True
Land.MenuSfx(CFGSound).Checked = True
Land.MenuBonus(CFGBonusTroops).Checked = True
Land.MenuInitTroops(CFGInitTroopPl).Checked = True
Land.MenuConquer(CFGConquer).Checked = True
Land.MenuShips(CFGShips).Checked = True
Land.MenuInitTroopCts(CFGInitTroopCt).Checked = True
Land.MenuPorts(CFGPorts).Checked = True
Land.MenuRandom(CFGEvents).Checked = True
Land.Menu1st(CFG1st).Checked = True
Land.MenuHQSelect(CFGHQSelect).Checked = True
Land.MenuSize(CFGCountrySize).Checked = True
Land.MenuPct(CFGLandPct).Checked = True
Land.MenuLakeSize(CFGLakeSize).Checked = True
Land.MenuIslands(CFGIslands).Checked = True
Land.MenuShape(CFGShape).Checked = True
Land.MenuProp(CFGProportion).Checked = True
Land.MenuAISpeed(CFGAISpeed).Checked = True
Land.MenuResolution(CFGResolution).Checked = True

End Sub

Public Sub ClearMenuChecks()

Dim i As Integer

'This sub erases the checkmarks on all menu items.
For i = 0 To 6
  If i <= 2 And i > 0 Then
    Land.MenuHQSelect(i).Checked = False
    Land.MenuSfx(i).Checked = False
  End If
  If i <= 3 And i > 0 Then
    Land.MenuRandom(i).Checked = False
    Land.MenuInitTroopCts(i).Checked = False
    Land.MenuIslands(i).Checked = False
    Land.MenuShape(i).Checked = False
    Land.MenuProp(i).Checked = False
    Land.MenuResolution(i).Checked = False
  End If
  If i <= 4 And i > 0 Then
    Land.MenuPorts(i).Checked = False
    Land.MenuShips(i).Checked = False
    Land.MenuLakeSize(i).Checked = False
    Land.MenuAISpeed(i).Checked = False
  End If
  If i <= 5 And i > 0 Then
    Land.MenuInitTroops(i).Checked = False
    Land.MenuConquer(i).Checked = False
    Land.MenuPct(i).Checked = False
  End If
  If i <= 5 Then
    Land.MenuBonus(i).Checked = False
  End If
  If i <= 6 And i > 0 Then
    Land.MenuSize(i).Checked = False
    Land.MenuBorders(i).Checked = False
  End If
  If i <= 7 And i > 0 Then
    Land.Menu1st(i).Checked = False
  End If
Next i

End Sub

Public Sub WriteDebugInfo()

If Not IsInDebug Then
    Exit Sub
End If

'Debug junk.  Do this at the end of initializing for a new game to write
'a debug file into the app's directory.  Useful for reporting bugs.
MakeINI App.Path & "\DEBUG.map"
WriteMapData App.Path & "\DEBUG.map"

End Sub

Public Function IsInDebug() As Boolean
    'This function will return whether or not the program is running in the IDE
    
    On Error GoTo IDERunning
    
    'causes an error only in the IDE as Debug statements are removed during compile
    Debug.Assert 1 / 0
    
    IsInDebug = False
    Exit Function
    
IDERunning:
    IsInDebug = True
End Function
