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
Public NextTile As Tile
Public Heading As Direction
Public CommandedHeading As Direction
Public Speed As Double

Private Sub IGamePiece_Move()

    If NextTile.IsTraversable Then
        Set CurrentTile = NextTile
    Else
        '//Get pathing options
        '//Calculate shortest path to target tile
        Set CurrentTile = MapManager.ChooseShorterPathTile(TargetTile, GetNextTileOptions())
    End If
    
    '//Aquire a the next desired tile and the next tile along the current path
    Set TargetTile = mActiveBehavior.AquireTargetTile(Nothing)
    Set Me.NextTile = MapManager.GetNextTile(CurrentTile, Heading)
End Sub

Private Function GetNextTileOptions() As Variant
    
End Function
