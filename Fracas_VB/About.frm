VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Fracas..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Jason Merlo..."
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Fracas..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Landmass..."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub Command1_Click()

About.Hide

End Sub

Private Sub Form_Load()

Label1.Caption = VersionStr
Label2.Caption = ""
Label3.Caption = WebStr & vbCrLf & EmailStr

End Sub

