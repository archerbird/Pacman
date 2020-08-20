VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine")
Option Explicit

Private centerCell As Range

Public HasDot As Boolean
Public HasPowerUp As Boolean
Public Occupant As IGamePiece

Public LeftPathNode As PathUnit
Public RightPathNode As PathUnit
Public UpperPathNode As PathUnit
Public LowerPathNode As PathUnit

