VERSION 5.00
Begin VB.Form DebugForm 
   Caption         =   "Debugger"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   Icon            =   "DebugForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Label XMyClientIndex 
      Height          =   255
      Left            =   720
      TabIndex        =   96
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   $"DebugForm.frx":0442
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
      TabIndex        =   95
      Top             =   2760
      Width           =   10815
   End
   Begin VB.Label Label5 
      Caption         =   "MsgQs"
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
      TabIndex        =   94
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   9
      Left            =   9120
      TabIndex        =   93
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   9
      Left            =   7320
      TabIndex        =   92
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   9
      Left            =   5520
      TabIndex        =   91
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   90
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   89
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   88
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   8
      Left            =   9120
      TabIndex        =   87
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   86
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   85
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   8
      Left            =   3720
      TabIndex        =   84
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   83
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   82
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   7
      Left            =   9120
      TabIndex        =   81
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   7
      Left            =   7320
      TabIndex        =   80
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   79
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   78
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   77
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   76
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   6
      Left            =   9120
      TabIndex        =   75
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   6
      Left            =   7320
      TabIndex        =   74
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   73
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   72
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   6
      Left            =   1920
      TabIndex        =   71
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   70
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   5
      Left            =   9120
      TabIndex        =   69
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   5
      Left            =   7320
      TabIndex        =   68
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   67
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   66
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   65
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   64
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   4
      Left            =   9120
      TabIndex        =   63
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   62
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   61
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   60
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   59
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   58
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   3
      Left            =   9120
      TabIndex        =   57
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   56
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   55
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   54
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   53
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   52
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   2
      Left            =   9120
      TabIndex        =   51
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   50
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   49
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   48
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   47
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   45
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   44
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   43
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   42
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   41
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label XMsg6 
      Height          =   255
      Index           =   10
      Left            =   9120
      TabIndex        =   39
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XMsg5 
      Height          =   255
      Index           =   10
      Left            =   7320
      TabIndex        =   38
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XMsg4 
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   37
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XMsg3 
      Height          =   255
      Index           =   10
      Left            =   3720
      TabIndex        =   36
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XMsg2 
      Height          =   255
      Index           =   10
      Left            =   1920
      TabIndex        =   35
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XMsg1 
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label XSeqNumber 
      Height          =   255
      Left            =   7080
      TabIndex        =   33
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Ssss 
      Caption         =   "Seq"
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
      Left            =   6600
      TabIndex        =   32
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Phase"
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
      Left            =   4920
      TabIndex        =   31
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Turn"
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
      Left            =   3360
      TabIndex        =   30
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "TIP"
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
      Left            =   1680
      TabIndex        =   29
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Xnb 
      Caption         =   "MyCI"
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
      TabIndex        =   28
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label XPhase 
      Height          =   255
      Left            =   5640
      TabIndex        =   27
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label XTurn 
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label XTurnInProgress 
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Net       Dat       Qtim      StringQ"
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
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   23
      Top             =   1320
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   21
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Top             =   1080
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Top             =   840
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Top             =   840
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.Label XStringQ 
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   8655
   End
   Begin VB.Label XQtimer 
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label XDataArray 
      Height          =   255
      Index           =   6
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label XNetArray 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
