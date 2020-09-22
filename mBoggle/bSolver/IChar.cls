VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrChar                As String
Private mintIndex               As Integer
Private mintGridX               As Integer
Private mintGridY               As Integer

Friend Property Let Char(ByVal strChar As String)
    mstrChar = strChar
End Property

Public Property Get Char() As String
Attribute Char.VB_UserMemId = 0
    Char = mstrChar
End Property

Friend Property Let Index(ByVal intIndex As Integer)
    mintIndex = intIndex
End Property

Public Property Get Index() As Integer
    Index = mintIndex
End Property

Friend Property Let GridX(ByVal intGridX As Integer)
    mintGridX = intGridX
End Property

Public Property Get GridX() As Integer
    GridX = mintGridX
End Property

Friend Property Let GridY(ByVal intGridY As Integer)
    mintGridY = intGridY
End Property

Public Property Get GridY() As Integer
    GridY = mintGridY
End Property
