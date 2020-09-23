VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private cXpos As Single
Private cYpos As Single
Private cXvel As Single
Private cYvel As Single
Private cYstart As Single
Private cYtilt As Single
Private cYtiltvel As Single
Private cYshadow As Single
Private cElastic As Single
Private cShape As Integer
Private cColor As Integer
Private cEnabled As Boolean

Property Get Xpos() As Single
Xpos = cXpos
End Property

Property Get Ypos() As Single
Ypos = cYpos
End Property

Property Get Xvel() As Single
Xvel = cXvel
End Property

Property Get Yvel() As Single
Yvel = cYvel
End Property

Property Get Ystart() As Single
Ystart = cYstart
End Property

Property Get Ytilt() As Single
Ytilt = cYtilt
End Property

Property Get Ytiltvel() As Single
Ytiltvel = cYtiltvel
End Property

Property Get Yshadow() As Single
Yshadow = cYshadow
End Property

Property Get Elastic() As Single
Elastic = cElastic
End Property

Property Get Shape() As Integer
Shape = cShape
End Property

Property Get Color() As Integer
Color = cColor
End Property

Property Get Enabled() As Boolean
Enabled = cEnabled
End Property

Property Let Xpos(x As Single)
cXpos = x
End Property

Property Let Ypos(x As Single)
cYpos = x
End Property

Property Let Xvel(x As Single)
cXvel = x
End Property

Property Let Yvel(x As Single)
cYvel = x
End Property

Property Let Ystart(x As Single)
cYstart = x
End Property

Property Let Ytilt(x As Single)
cYtilt = x
End Property

Property Let Ytiltvel(x As Single)
cYtiltvel = x
End Property

Property Let Yshadow(x As Single)
cYshadow = x
End Property

Property Let Elastic(x As Single)
cElastic = x
End Property

Property Let Shape(x As Integer)
cShape = x
End Property

Property Let Color(x As Integer)
cColor = x
End Property

Property Let Enabled(y As Boolean)
cEnabled = y
End Property
