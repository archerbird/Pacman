VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ghost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.Model.Ghost"
Option Explicit
Implements IGamePiece

Private mActiveBehavior As IGhostBehavior
Public TargetTile As Tile
Public CurrentTile As Tile
Public Heading As Direction
Public Speed As Double

Private Sub IGamePiece_Move()
    '//Go to target tile
    Set CurrentTile = TargetTile
    '//Aquire a new target tile
    Set TargetTile = mActiveBehavior.AquireTargetTile(Nothing)
End Sub
