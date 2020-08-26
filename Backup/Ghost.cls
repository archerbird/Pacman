VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ghost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.GamePiece.Ghost")
Option Explicit
Implements IGamePiece

Private mActiveBehavior As IGhostBehavior
Private mTargetTile As Tile
Private mCurrentTile As Tile

Private Sub IGamePiece_Move()
    '//Go to target tile
    Set mCurrentTile = mTargetTile
    '//Aquire a new target tile
    Set mTargetTile = mActiveBehavior.AquireTargetTile(Nothing)
End Sub
