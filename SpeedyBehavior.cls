VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpeedyBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Ghost.PathingAI"

Option Explicit
'@PredeclaredId

Implements IGhostBehavior

Private Type TSpeedyBehavior
    Target As PacmanModel
End Type

Private this As TSpeedyBehavior

Public Function Create(pm As PacmanModel) As IGhostBehavior
    Dim obj As New ShadowBehavior
    Set obj.Target = pm
    Set Create = obj
End Function

'//Targets 4 tiles ahead of pacman
Private Function IGhostBehavior_AquireTargetTile() As Tile
    Set IGhostBehavior_AquireTargetTile = MapManager.GetNextTile(this.Target.CurrentTile, this.Target.Heading, 8)
End Function

Public Property Set Target(value As PacmanModel)
    Set this.Target = value
End Property

