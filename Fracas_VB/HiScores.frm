VERSION 5.00
Begin VB.Form HiScores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hi Scores"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4065
   Icon            =   "HiScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OkBootie 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   22
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   240
      TabIndex        =   33
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "10."
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
      Index           =   8
      Left            =   240
      TabIndex        =   32
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "9."
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
      Index           =   7
      Left            =   240
      TabIndex        =   31
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "8."
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
      Index           =   6
      Left            =   240
      TabIndex        =   30
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "7."
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
      Index           =   5
      Left            =   240
      TabIndex        =   29
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "6."
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
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "5."
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
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "4."
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
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "3."
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
      Index           =   1
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "2."
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
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Ranklbl 
      Alignment       =   2  'Center
      Caption         =   "Rank"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   495
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   9
      Left            =   3120
      TabIndex        =   21
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   8
      Left            =   3120
      TabIndex        =   20
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   7
      Left            =   3120
      TabIndex        =   19
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   6
      Left            =   3120
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   5
      Left            =   3120
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   4
      Left            =   3120
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
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
      Index           =   10
      Left            =   3120
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   10
      Left            =   840
      TabIndex        =   12
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   9
      Left            =   840
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   8
      Left            =   840
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   7
      Left            =   840
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   6
      Left            =   840
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   5
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
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
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Scorelbl 
      Alignment       =   2  'Center
      Caption         =   "Score"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label NameLbl 
      Alignment       =   2  'Center
      Caption         =   "Name"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label HiScorez 
      Alignment       =   2  'Center
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label HiNamez 
      Alignment       =   2  'Center
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "HiScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Public Sub ClearHiScores()

'This sub simply sets up the 'default' hi score list.
Dim DefaultName(10) As String
Dim Order(10) As Integer
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim k As Integer

DefaultName(1) = "Bowie"
DefaultName(2) = "BooBoo"
DefaultName(3) = "Pinky"
DefaultName(4) = "Zack"
DefaultName(5) = "Smudge"
DefaultName(6) = "Ozzie"
DefaultName(7) = "Mac"
DefaultName(8) = "Penny"
DefaultName(9) = "Lady"
DefaultName(10) = "Queenie"

'Set up the order array...
For i = 1 To 10: Order(i) = i: Next
'And randomize it so 1-10 appear in a random order.
For i = 1 To 20
  Do
    j = Int(Rnd(1) * 10) + 1
    l = Int(Rnd(1) * 10) + 1
  Loop Until (j <> l)
  k = Order(j)
  Order(j) = Order(l)
  Order(l) = k
Next i

'Now set up the hi scores!
For i = 1 To 10
  HiScoreName(i) = DefaultName(Order(i))
  HiScore(i) = ((10 - i) * 10) + (Int(Rnd(1) * 9) + 1)
  HiScoreColor(i) = 0
Next i

End Sub

Public Sub AddPlayerToHighScoreList(LuckyGuy As Integer)

Dim i As Integer
Dim j As Integer
Dim temp As Long
'This sub adds the passed player's name and score to the high score list.
'If they don't belong there, this sub does nothing.
For i = 1 To 10
  If STATscore(LuckyGuy) > HiScore(i) Then
    'The player belongs at position i.
    'First, move everyone from here down a notch.
    If i < 10 Then  'Don't bother if it's the last one.
      For j = 9 To i Step -1
        HiScore(j + 1) = HiScore(j)
        HiScoreName(j + 1) = HiScoreName(j)
        HiScoreColor(j + 1) = HiScoreColor(j)
      Next j
    End If
    'Now put us in its place.
    HiScore(i) = STATscore(LuckyGuy)
    HiScoreName(i) = PlayerName(LuckyGuy)
    HiScoreColor(i) = Player(LuckyGuy)
    Exit For
  End If
Next i

End Sub

Private Sub Form_Load()

'Someone wants to view the high score list!
'We only get here during a game, so our scores *must* exist.

Dim i As Integer

For i = 1 To NUM_HI_SCORES
  HiNamez(i).Caption = HiScoreName(i)
  HiScorez(i).Caption = Trim(Str(HiScore(i)))
  If HiScoreColor(i) = 0 Then
    HiNamez(i).BackColor = HiScores.BackColor
  Else
    HiNamez(i).BackColor = PlayerColorCodes(HiScoreColor(i))
  End If
Next i

End Sub

Private Sub OkBootie_Click()

Unload Me

End Sub
