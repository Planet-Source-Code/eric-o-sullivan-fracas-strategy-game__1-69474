VERSION 5.00
Begin VB.Form ChatForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chat"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9510
   Icon            =   "ChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ExitBut 
      Caption         =   "Exit"
      Height          =   285
      Left            =   8520
      TabIndex        =   26
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton SendBut 
      Caption         =   "Send"
      Height          =   285
      Left            =   7560
      TabIndex        =   25
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox ChatText 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   3120
      Width           =   7335
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   11
      Left            =   1560
      TabIndex        =   23
      Top             =   2520
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   10
      Left            =   1560
      TabIndex        =   21
      Top             =   2280
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   9
      Left            =   1560
      TabIndex        =   19
      Top             =   2040
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   8
      Left            =   1560
      TabIndex        =   17
      Top             =   1800
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   7
      Left            =   1560
      TabIndex        =   15
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   6
      Left            =   1560
      TabIndex        =   13
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   5
      Left            =   1560
      TabIndex        =   11
      Top             =   1080
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   4
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label TextSlot 
      Caption         =   "Label1"
      Height          =   225
      Index           =   12
      Left            =   1560
      TabIndex        =   1
      Top             =   2760
      Width           =   7815
   End
   Begin VB.Label NameSlot 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   225
      Index           =   12
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "ChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const NUM_CHAT_LINES = 12

Private Sub Form_Load()

Dim i As Integer

For i = 1 To NUM_CHAT_LINES
  NameSlot(i).Caption = vbNullString
  TextSlot(i).Caption = vbNullString
  NameSlot(i).BackColor = ChatForm.BackColor
  TextSlot(i).BackColor = ChatForm.BackColor
  ChatText.Text = vbNullString
  SendBut.Default = True
Next i

End Sub

Private Sub ExitBut_Click()

Me.Hide

End Sub

Private Sub SendBut_Click()

Dim Sender As Integer
Dim HumanCt As Integer
Dim i As Integer

'Leave if not in a network game
If MyNetworkRole = NW_NONE Then Exit Sub

'We are in a network game.  Send this text on out.
If ChatText.Text = vbNullString Then Exit Sub

If MyClientIndex = 0 Then
  'We're a server, and we could have *several* humans on this machine chatting.
  'If there's only one human, then that's the sender.  Otherwise, the sender is HOST.
  HumanCt = 0
  For i = 1 To MAX_PLAYERS
    If PlayerType(i) = PTYPE_HUMAN Then
      HumanCt = HumanCt + 1
      Sender = i
    End If
  Next i
  If HumanCt <> 1 Then Sender = 0
Else
  'We're a client, the sender is us.
  Sender = MyClientIndex
End If

'Send it to the network.
SendChatToNetwork Sender, RemoveLeadingSpace(ChatText.Text), 0
'Update our own chat lines.
ChatTextReceived Sender, RemoveLeadingSpace(ChatText.Text)

ChatText.Text = vbNullString
ChatText.SetFocus

End Sub

Public Sub ChatTextReceived(PlayerNum As Integer, TextLine As String)

Dim ChName As String
Dim ChBackColor As Long
Dim ChForeColor As Long
Dim TxBackColor As Long
Dim TxForeColor As Long
Dim TxBold As Boolean

BumpUpChatLines

If PlayerNum = 0 Then
  ChName = "Host"
  ChBackColor = vbWhite
  ChForeColor = vbBlack
  TxBackColor = ChatForm.BackColor
  TxForeColor = vbBlack
Else
  ChName = PlayerName(PlayerNum)
  ChBackColor = PlayerColorCodes(Player(PlayerNum))
  ChForeColor = PlayerTextColor(Player(PlayerNum))
  TxBackColor = ChatForm.BackColor
  TxForeColor = vbBlack
End If

TxBold = False
If LCase(Left(TextLine, 3)) = "/me" Then
  TextLine = Right(TextLine, Len(TextLine) - 3)
  TextLine = RemoveLeadingSpace(TextLine)
  If TextLine = vbNullString Then TextLine = "..."
  TextLine = ChName & " " & TextLine
  TxBold = True
  TxBackColor = ChBackColor
  ChName = vbNullString
  ChBackColor = ChatForm.BackColor
End If

NameSlot(NUM_CHAT_LINES).Caption = ChName
NameSlot(NUM_CHAT_LINES).BackColor = ChBackColor
NameSlot(NUM_CHAT_LINES).ForeColor = ChForeColor
TextSlot(NUM_CHAT_LINES).FontBold = TxBold
TextSlot(NUM_CHAT_LINES).Caption = TextLine
TextSlot(NUM_CHAT_LINES).BackColor = TxBackColor
TextSlot(NUM_CHAT_LINES).ForeColor = TxForeColor

End Sub

Private Sub BumpUpChatLines()

Dim i As Integer

'This sub just moves everything up one slot to make room for a new chat line.

For i = 2 To NUM_CHAT_LINES
  NameSlot(i - 1).Caption = NameSlot(i).Caption
  NameSlot(i - 1).BackColor = NameSlot(i).BackColor
  NameSlot(i - 1).ForeColor = NameSlot(i).ForeColor
  TextSlot(i - 1).FontBold = TextSlot(i).FontBold
  TextSlot(i - 1).Caption = TextSlot(i).Caption
  TextSlot(i - 1).BackColor = TextSlot(i).BackColor
  TextSlot(i - 1).ForeColor = TextSlot(i).ForeColor
Next i

End Sub

Private Function RemoveLeadingSpace(MyStr As String)

Dim Done As Boolean

Done = False
Do
  If Left(MyStr, 1) = " " Then
    MyStr = Right(MyStr, Len(MyStr) - 1)
  Else
    Done = True
  End If
Loop Until Done = True

RemoveLeadingSpace = MyStr

End Function
