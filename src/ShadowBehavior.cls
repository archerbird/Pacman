VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShadowBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PacmanGame.Model.Ghost.PathingAI"

Option Explicit
Implements IGhostBehavior

Private Type TShadowBehavior
    Target As PacmanModel
End Type

Private this As TShadowBehavior

Public Function Create(pm As PacmanModel) As IGhostBehavior
    Dim obj As New ShadowBehavior
    Set obj.Target = pm
    Set Create = obj
End Function

Private Function IGhostBehavior_AquireTargetTile() As Tile
    Set IGhostBehavior_AquireTargetTile = this.Target.CurrentTile
End Function

Public Property Set Target(value As PacmanModel)
    Set this.Target = value
End Property
