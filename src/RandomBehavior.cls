VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RandomBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Ghost.PathingAI"
'@PredeclaredId
Option Explicit

Implements IGhostBehavior


Public Function Create() As IGhostBehavior
    Dim obj As New RandomBehavior
    Set Create = obj
End Function

Private Function IGhostBehavior_AquireTargetTile() As Tile
    Dim offsetTile As Tile
    Dim randomNumber As Integer
    
    Set IGhostBehavior_AquireTargetTile = MapManager.GetMazeTile( _
                                                         Int(60 * Rnd), _
                                                         Int(60 * Rnd) _
                                                        )
                                                        
End Function


