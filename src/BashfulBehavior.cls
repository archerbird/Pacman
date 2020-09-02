VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BashfulBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Ghost.PathingAI"

Option Explicit
'@PredeclaredId
Implements IGhostBehavior

Private Type TBashfulBehavior
    blinky As GhostModel
    Pacman As PacmanModel
End Type

Private this As TBashfulBehavior

Public Function Create(pm As PacmanModel, bl As GhostModel) As IGhostBehavior
    Dim obj As New BashfulBehavior
    Set obj.Pacman = pm
    Set obj.blinky = bl
    Set Create = obj
End Function

Public Property Set blinky(ByVal RHS As GhostModel)
    Set this.blinky = RHS
End Property


Public Property Set Pacman(ByVal RHS As PacmanModel)
    Set this.Pacman = RHS
End Property

Private Function IGhostBehavior_AquireTargetTile() As Tile
    Dim offsetTile As Tile
    
    Set offsetTile = MapManager.GetNextTile(this.Pacman.CurrentTile, this.Pacman.Heading, 3)
    
    Set IGhostBehavior_AquireTargetTile = MapManager.GetMazeTile( _
                                                        this.blinky.CurrentTile.y + 2 * (offsetTile.y - this.blinky.CurrentTile.y), _
                                                        this.blinky.CurrentTile.x + 2 * (offsetTile.x - this.blinky.CurrentTile.x) _
                                                        )
                                                        
End Function


