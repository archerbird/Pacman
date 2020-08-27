VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScatterBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.Model.Ghost.PathingAI")
Option Explicit
Implements IGhostBehavior
Public CornerTile As Tile


Private Function IGhostBehavior_AquireTargetTile() As Tile
    Set IGhostBehavior_AquireTargetTile = MapManager.GetMazeTile(1, 1)
End Function
