VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private cXpos As Long
Private cYpos As Long
Private cSpeed As Single
Private cCount As Single
Private cFrame As Integer
Private cShape As Integer
Private cEnabled As Boolean

Property Get Xpos() As Long
Xpos = cXpos
End Property

Property Get Ypos() As Long
Ypos = cYpos
End Property

Property Get Speed() As Single
Speed = cSpeed
End Property

Property Get Count() As Single
Count = cCount
End Property

Property Get Frame() As Integer
Frame = cFrame
End Property

Property Get Shape() As Integer
Shape = cShape
End Property

Property Get Enabled() As Boolean
Enabled = cEnabled
End Property

Property Let Xpos(x As Long)
cXpos = x
End Property

Property Let Ypos(x As Long)
cYpos = x
End Property

Property Let Speed(x As Single)
cSpeed = x
End Property

Property Let Count(x As Single)
cCount = x
End Property

Property Let Frame(x As Integer)
cFrame = x
End Property

Property Let Shape(x As Integer)
cShape = x
End Property

Property Let Enabled(y As Boolean)
cEnabled = y
End Property
