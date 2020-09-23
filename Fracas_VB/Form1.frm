VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox NumTroops 
      Height          =   285
      Left            =   5880
      TabIndex        =   17
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go !"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox Dprop1 
      Height          =   285
      Left            =   5880
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Aprop1 
      Height          =   285
      Left            =   5880
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Dprop2 
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Aprop2 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox DSum1 
      Height          =   285
      Left            =   4440
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox DSum2 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox ASum2 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox ASum1 
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox A3t 
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox A2t 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox A1t 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox B3t 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox B2t 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox B1t 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox D1t 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox D2t 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

A1 = Val(A1t.Text)
A2 = Val(A2t.Text)
A3 = Val(A3t.Text)
D1 = Val(D1t.Text)
D2 = Val(D2t.Text)
B1 = Val(B1t.Text)
B2 = Val(B2t.Text)
B3 = Val(B3t.Text)

attsum1 = (A1 / (A1 + D1)) + (A2 / (A2 + D1)) + (A3 / (A3 + D1))
attsum2 = (B1 / (B1 + D2)) + (B2 / (B2 + D2)) + (B3 / (B3 + D2))
defsum1 = (D1 / (A1 + D1)) + (D1 / (A2 + D1)) + (D1 / (A3 + D1))
defsum2 = (D2 / (B1 + D2)) + (D2 / (B2 + D2)) + (D2 / (B3 + D2))

attprop1 = attsum1 / (attsum1 + attsum2)
attprop2 = attsum2 / (attsum1 + attsum2)
defprop1 = defsum1 / (defsum1 + defsum2)
defprop2 = defsum2 / (defsum1 + defsum2)

ASum1.Text = Trim(Str(attsum1))
ASum2.Text = Trim(Str(attsum2))
DSum1.Text = Trim(Str(defsum1))
DSum2.Text = Trim(Str(defsum2))

Aprop1.Text = Trim(Str(attprop1))
Aprop2.Text = Trim(Str(attprop2))
Dprop1.Text = Trim(Str(defprop1))
Dprop2.Text = Trim(Str(defprop2))

If Aprop1.Text < 0.5 Then
  NumTroops.Text = "0"
Else
  troopcalc = Aprop1 - 0.5
  troopcalc = troopcalc * 2 * D2
  NumTroops.Text = Trim(Str(troopcalc))
End If

End Sub
