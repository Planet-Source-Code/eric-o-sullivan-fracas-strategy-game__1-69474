VERSION 5.00
Begin VB.Form RenameCountry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rename"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   4320
   Icon            =   "RenameCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButt 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CancelButt 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox NewName 
      Height          =   285
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "New Name"
      Top             =   555
      Width           =   2055
   End
   Begin VB.Label OldName 
      Caption         =   "Old Name"
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
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "New Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Old Name:"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "RenameCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub CancelButt_Click()

Me.Hide

End Sub

Private Sub OKbutt_Click()

'See if our name is valid.
If NewName.Text <> vbNullString Then
  'Assign this name to our country!
  If LastRightClick > TILEVAL_COASTLINE Then
    MyMap.WaterName(LastRightClick - 1000) = NewName.Text
  Else
    MyMap.CountryName(LastRightClick) = NewName.Text
  End If
  Land.CntryName = NewName.Text
  'Also send this over the network.
  If MyNetworkRole <> NW_NONE Then
    SendRenameToNetwork LastRightClick, NewName.Text, 0
  End If
  Me.Hide
End If

End Sub
