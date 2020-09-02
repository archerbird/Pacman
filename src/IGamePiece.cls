VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGamePiece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model"
'@Interface
Option Explicit

Public Property Get Id() As String
End Property

Public Property Get CycleRemainder() As Integer
End Property
Public Property Let CycleRemainder(value As Integer)
End Property

Public Property Get Speed() As Double
End Property

Public Property Get CurrentHeading() As Direction
End Property

Public Property Get CurrentTile() As Tile
End Property

Public Property Get TargetTile() As Tile
End Property

Public Sub Move()
End Sub
