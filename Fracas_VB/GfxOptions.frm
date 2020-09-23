VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form GfxOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graphics Options"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "GfxOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox PromptCheckbox 
      Caption         =   "Prompt to Save Map"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox FlashingCheckbox 
      Caption         =   "Flash Countries"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   720
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox UnOccColor 
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1500
      Width           =   1455
   End
   Begin VB.CheckBox WavesCheckbox 
      Caption         =   "Enable Waves"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox ExplosionsCheckbox 
      Caption         =   "Enable Explosions"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton CANCELbutton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton OKbutton 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin MSComctlLib.Slider AniSpeed 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   200
      SelStart        =   20
      TickFrequency   =   20
      Value           =   20
   End
   Begin VB.Label Label4 
      Caption         =   "Unoccupied Country Color:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Overall Animation Speed"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Slow"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Fast"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "GfxOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Dim TempSpeed As Long

Private Sub AniSpeed_Change()

Land.BallTimer.Interval = AniSpeed.Value

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then
  CANCELbutton_Click
End If

End Sub

Private Sub OKbutton_Click()

Dim i As Long
Dim j As Integer

'OK!  Take these settings and record them.

'AniSpeed is already set.

'If we disabled explosions, clean out the balls array.
CFGExplosions = ExplosionsCheckbox.Value
If CFGExplosions = 0 Then ClearAllBalls

'If we disabled waves, clean out the waves array.
CFGWaves = WavesCheckbox.Value
If CFGWaves = 0 Then ClearAllWaves

'If we disabled flashing, return each country to normal state.
CFGFlashing = FlashingCheckbox.Value
If (CFGFlashing = 0) And (GameMode = GM_GAME_ACTIVE) Then ResetFlashing

'Get the prompt check.
CFGPrompt = PromptCheckbox.Value

'Grab the Unoccupied Country color.
For j = 1 To 11
  If UnOccColor.Text = PlayerColors(j) Then
    'We have chosen color j for unoccupied countries.
    CFGUnoccupiedColor = j
  End If
Next j
'Set all unoccupied countries to this color if a game is in progress.
If GameMode = GM_GAME_ACTIVE Then
  For i = 1 To MyMap.NumberOfCountries
    If MyMap.Owner(i) = 0 Then
      'Aha!  This one is unowned.
      MyMap.CountryColor(i) = CFGUnoccupiedColor
    End If
  Next i
  DrawMap   'Update the screen immediately.
ElseIf (GameMode = GM_TITLE_DIALOG_OPEN) Then
  'Title screen.  Change the color of the Fracas word to the unoccupied color.
  MyMap.CountryColor(1) = CFGUnoccupiedColor
  DrawMap
End If

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

Unload Me

End Sub

Private Sub CANCELbutton_Click()

'CANCEL!  Put the settings back where they were.
Land.BallTimer.Interval = TempSpeed

'Clean out the balls and waves arrays.
ClearAllBalls
ClearAllWaves

Unload Me

End Sub

Private Sub Form_Load()

Dim j As Integer
Dim k As Integer
Dim FoundIt As Boolean

'Set up initial positions of controls.
TempSpeed = Land.BallTimer.Interval
AniSpeed.Value = TempSpeed
ExplosionsCheckbox.Value = CFGExplosions
WavesCheckbox.Value = CFGWaves
FlashingCheckbox.Value = CFGFlashing
PromptCheckbox.Value = CFGPrompt

'Set up color dropdown.
For j = 1 To 11       'Not 12 since white is reserved for flashing.
  'Don't add player colors to the dropdown.
  FoundIt = False
  'See if this is one of the player colors.
  For k = 1 To 6
    If (Player(k) = j) Then FoundIt = True
  Next k
  If FoundIt = False Then UnOccColor.AddItem PlayerColors(j)
Next j
UnOccColor.Text = PlayerColors(CFGUnoccupiedColor)
UpdateColorDropdown

End Sub

Private Sub UpdateColorDropdown()

'This sub just writes the chosen color to the BackColor of the
'color drop-down.

Dim j As Integer

For j = 1 To 12
  If UnOccColor.Text = PlayerColors(j) Then
    UnOccColor.BackColor = PlayerColorCodes(j)
    UnOccColor.ForeColor = PlayerTextColor(j)
  End If
Next j

End Sub

Private Sub UnOccColor_Click()

UpdateColorDropdown
If GfxOptions.Visible = True Then
  OKbutton.SetFocus
End If

End Sub
