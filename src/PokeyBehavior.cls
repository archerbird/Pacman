VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PokeyBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("PacmanGame.Model.Ghost.PathingAI")

Option Explicit
Implements IGhostBehavior

Private Type TPokeyBehavior
    Player As PacmanModel
    GhostSelf As GhostModel
    ScatterTarget As Tile
End Type

Private this As TPokeyBehavior

Public Function Create(pm As PacmanModel, self As GhostModel, scatterTile As Tile) As IGhostBehavior
    Dim obj As New PokeyBehavior
    obj.Init pm, self, scatterTile
    Set Create = obj
End Function

Public Function Init(pm As PacmanModel, self As GhostModel, scatterTile As Tile)
    Set this.Player = pm
    Set this.GhostSelf = self
    Set this.ScatterTarget = scatterTile
End Function

'//Targets pacman only if we are more than 8 tiles away. Otherwise, target the scartter tile.
Private Function IGhostBehavior_AquireTargetTile() As Tile
    If MapManager.TileDistance(this.Player.CurrentTile, this.GhostSelf.CurrentTile) > 13 Then
        Set IGhostBehavior_AquireTargetTile = MapManager.GetNextTile(this.Player.CurrentTile, this.Player.Heading)
    Else
        Set IGhostBehavior_AquireTargetTile = this.ScatterTarget
    End If
End Function


