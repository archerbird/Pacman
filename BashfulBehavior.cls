VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BashfulBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("PacManEngine.Model.Ghost.PathingAI")

Option Explicit
'@PredeclaredId
Implements IGhostBehavior

Private Type TBashfulBehavior
    blinky As Ghost
    Pacman As PacmanController
End Type

Private this As TBashfulBehavior

Public Function Create(pm As PacmanController, bl As Ghost) As IGhostBehavior
    Dim obj As New BashfulBehavior
    Set obj.Pacman = pm
    Set obj.blinky = bl
    Set Create = obj
End Function

Public Property Set blinky(ByVal RHS As Ghost)
    Set this.blinky = RHS
End Property


Public Property Set Pacman(ByVal RHS As PacmanController)
    Set this.Pacman = RHS
End Property

Private Function IGhostBehavior_AquireTargetTile() As Tile
    Dim offsetTile As Tile
    
    Set offsetTile = MapManager.GetNextTile(this.Pacman.CurrentTile, this.Pacman.Heading, 2)
    
    Set IGhostBehavior_AquireTargetTile = MapManager.GetMazeTile( _
                                                        2 * Abs(offsetTile.y - this.Pacman.CurrentTile.y), _
                                                        2 * Abs(offsetTile.x - this.Pacman.CurrentTile.x) _
                                                        )
                                                        
End Function
