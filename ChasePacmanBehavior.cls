VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChasePacmanBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("PacManEngine.Model.Ghost.PathingAI")

Option Explicit
Implements IGhostBehavior

Private Type TChasePacmanBehavior
    Target As PacmanController
End Type

Private this As TChasePacmanBehavior

Public Function Create(pm As PacmanController) As IGhostBehavior
    Dim obj As New ChasePacmanBehavior
    Set obj.Target = pm
    Set Create = obj
End Function

Private Function IGhostBehavior_AquireTargetTile() As Tile
    Set IGhostBehavior_AquireTargetTile = this.Target.CurrentTile
End Function

Public Property Set Target(value As PacmanController)
    Set this.Target = value
End Property
