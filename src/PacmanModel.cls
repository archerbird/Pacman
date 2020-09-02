VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PacmanModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Pacman"
Option Explicit
Implements IGamePiece

Private Const CHARACTER_ID As String = "pacman"

Private Type TPacmanController
    NextTile As Tile '// the next tile in pacman's current path
    CurrentTile As Tile '// the tile pacman is current in
    TargetTile As Tile '// the tile that pacman wants to go to
    Heading As Direction '// the direction pacman is currently facing
    CommandedHeading As Direction '//the direction pacman will try to go
    Speed As Double
    CycleRemainder As Integer
    IsDead As Boolean
End Type

Private this As TPacmanController

Public Property Get Name() As String
    Name = CHARACTER_ID
End Property

Public Property Get NextTile() As Tile
    Set NextTile = this.NextTile
End Property

Public Property Set NextTile(ByVal RHS As Tile)
    Set this.NextTile = RHS
End Property

Public Property Get TargetTile() As Tile
    Set TargetTile = this.TargetTile
End Property

Public Property Set TargetTile(ByVal RHS As Tile)
    Set this.TargetTile = RHS
End Property

Public Property Get CurrentTile() As Tile
    Set CurrentTile = this.CurrentTile
End Property

Public Property Set CurrentTile(ByVal RHS As Tile)
    Set this.CurrentTile = RHS
    Set this.NextTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
    Set this.TargetTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
End Property

Public Property Get Heading() As Direction
    Heading = this.Heading
   
End Property

Public Property Let Heading(ByVal RHS As Direction)
    this.CommandedHeading = RHS
    If Not this.CurrentTile Is Nothing Then
        Set this.TargetTile = MapManager.GetNextTile(this.CurrentTile, this.CommandedHeading)
    End If
End Property

Public Property Get IsDead() As Boolean
    IsDead = this.IsDead
End Property

Public Property Let IsDead(ByVal RHS As Boolean)
    this.IsDead = RHS
End Property

Private Sub Class_Initialize()
    this.Speed = 0.8
    this.Heading = dLeft
    this.CommandedHeading = dLeft
End Sub


'//Movebale Gamepeice Implementation
Private Property Get IGamePiece_CurrentHeading() As Direction
    IGamePiece_CurrentHeading = Me.Heading
End Property

Private Property Get IGamePiece_CurrentTile() As Tile
    Set IGamePiece_CurrentTile = Me.CurrentTile
End Property

Private Property Let IGamePiece_CycleRemainder(RHS As Integer)
    this.CycleRemainder = RHS
End Property

Private Property Get IGamePiece_CycleRemainder() As Integer
    IGamePiece_CycleRemainder = this.CycleRemainder
End Property

Private Property Get IGamePiece_Id() As String
    IGamePiece_Id = CHARACTER_ID
End Property

Private Sub IGamePiece_Move()
    
    If this.TargetTile.IsTraversable Then
        '//Go to target tile
        this.Heading = this.CommandedHeading
        Set this.CurrentTile = this.TargetTile
        
    ElseIf this.NextTile.IsTraversable Then
        Set this.CurrentTile = this.NextTile
    Else
        '//can't move
        Me.IsDead = True
    End If
    
    '//Aquire a the next desired tile and the next tile along the current path
    Set Me.TargetTile = MapManager.GetNextTile(this.CurrentTile, this.CommandedHeading)
    Set Me.NextTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
End Sub

Private Property Get IGamePiece_Speed() As Double
    IGamePiece_Speed = this.Speed
End Property
