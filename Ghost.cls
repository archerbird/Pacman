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

Private Type TGhost
    ActiveBehavior As IGhostBehavior
    TargetTile As Tile
    CurrentTile As Tile
    NextTile As Tile
    Heading As Direction
    Speed As Double
End Type

Private this As TGhost

Public Property Set ActiveBehavior(ByVal RHS As IGhostBehavior)
    Set this.ActiveBehavior = RHS
    Set TargetTile = this.ActiveBehavior.AquireTargetTile()
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
End Property

Public Property Get NextTile() As Tile
    Set NextTile = this.NextTile
End Property

Public Property Set NextTile(ByVal RHS As Tile)
    Set this.NextTile = RHS
End Property

Public Property Get Heading() As Direction
    Heading = this.Heading
End Property

Public Property Let Heading(ByVal RHS As Direction)
    this.Heading = RHS
End Property

Public Property Get Speed() As Double
    Speed = this.Speed
End Property

Public Property Let Speed(ByVal RHS As Double)
    this.Speed = RHS
End Property

Public Sub Move()
    Call IGamePiece_Move
End Sub

Private Sub Class_Initialize()
    this.Heading = dLeft
    this.Speed = 0.7
End Sub

Private Sub IGamePiece_Move()

    If NextTile.IsTraversable Then
        Set CurrentTile = NextTile
    Else
        '//Get pathing options
        '//Calculate shortest path to target tile
        Set CurrentTile = MapManager.ChooseShorterPathTile(TargetTile, GetNextTileOptions())
    End If
    
    '//Aquire a the next desired tile and the next tile along the current path
    Set this.TargetTile = this.ActiveBehavior.AquireTargetTile()
    Set this.NextTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
End Sub

Private Function GetNextTileOptions() As Variant
    Dim TilesArr As Variant
    Dim nextIndex As Long
    Dim testTile As Tile
    
    If Heading <> dDown Then
        nextIndex = UBound(TilesArr) + 1
        Set testTile = MapManager.GetNextTile(CurrentTile, dDown)
        If testTile.IsTraversable Then
            ReDim Preserve TilesArr(LBound(TilesArr) To nextIndex)
            Set TilesArr(nextIndex) = testTile
        End If
    End If
    
    If Heading <> dUp Then
        nextIndex = UBound(TilesArr) + 1
        Set testTile = MapManager.GetNextTile(CurrentTile, dUp)
        If testTile.IsTraversable Then
            ReDim Preserve TilesArr(LBound(TilesArr) To nextIndex)
            Set TilesArr(nextIndex) = testTile
        End If
    End If
    
    If Heading <> dLeft Then
        nextIndex = UBound(TilesArr) + 1
        Set testTile = MapManager.GetNextTile(CurrentTile, dLeft)
        If testTile.IsTraversable Then
            ReDim Preserve TilesArr(LBound(TilesArr) To nextIndex)
            Set TilesArr(nextIndex) = testTile
        End If
    End If
    
    If Heading <> dRight Then
       nextIndex = UBound(TilesArr) + 1
        Set testTile = MapManager.GetNextTile(CurrentTile, dRight)
        If testTile.IsTraversable Then
            ReDim Preserve TilesArr(LBound(TilesArr) To nextIndex)
            Set TilesArr(nextIndex) = testTile
        End If
    End If
    
End Function
