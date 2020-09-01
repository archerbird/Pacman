VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GhostModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.Model.Ghost"
Option Explicit
Implements IGamePiece

Private Type TGhost
    ActiveBehavior As IGhostBehavior
    TargetTile As Tile
    CurrentTile As Tile
    NextTile As Tile
    Heading As Direction
    Speed As Double
    CycleRemainder As Integer
    Name As String
    Color As Long
End Type

Private this As TGhost

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(value As String)
    this.Name = value
End Property

Public Property Get Color() As Long
    Color = this.Color
End Property

Public Property Let Color(value As Long)
    this.Color = value
End Property

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
    Set this.NextTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
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

Private Sub Class_Initialize()
    this.Heading = dLeft
    this.Speed = 0.7
End Sub

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
    IGamePiece_Id = this.Name
End Property

Private Sub IGamePiece_Move()
    
    If CurrentTile.IsDecisionNode Or Not NextTile.IsTraversable Then
        '//Get pathing options
        '//Calculate shortest path to target tile
        ChooseNextTileAndDirection

    Else
        Set CurrentTile = NextTile
        
    End If
    
    If CurrentTile.Id = TileToken.SLOW_PATH_TOKEN Then
        this.Speed = 0.5
    Else
        this.Speed = 0.7
    End If
    
    '//Aquire a the next desired tile and the next tile along the current path
    Set this.TargetTile = this.ActiveBehavior.AquireTargetTile()
    Set this.NextTile = MapManager.GetNextTile(this.CurrentTile, this.Heading)
End Sub

Private Sub ChooseNextTileAndDirection()
    Dim DistToTarget As Long
    Dim testTile As Tile
    Dim selectedTile As Tile
    Dim minDist As Long
    Dim selectedDirection As Direction
    
    minDist = 2147483647 '//max long
    
    
    
    
    If Heading <> dDown Then
        Set testTile = MapManager.GetNextTile(CurrentTile, dUp)
        If testTile.IsTraversable Then
            DistToTarget = MapManager.TileDistance(this.TargetTile, testTile)
            If DistToTarget < minDist Then
                minDist = DistToTarget
                Set selectedTile = testTile
                selectedDirection = dUp
            End If
        End If
    End If
    
    If Heading <> dRight Then
        Set testTile = MapManager.GetNextTile(CurrentTile, dLeft)
        If testTile.IsTraversable Then
            DistToTarget = MapManager.TileDistance(this.TargetTile, testTile)
            If DistToTarget < minDist Then
                minDist = DistToTarget
                Set selectedTile = testTile
                selectedDirection = dLeft
            End If
        End If
    End If
    
    If Heading <> dUp Then
        Set testTile = MapManager.GetNextTile(CurrentTile, dDown)
        If testTile.IsTraversable Then
            DistToTarget = MapManager.TileDistance(this.TargetTile, testTile)
            If DistToTarget < minDist Then
                minDist = DistToTarget
                Set selectedTile = testTile
                selectedDirection = dDown
            End If
        End If
    End If
    
    If Heading <> dLeft Then
        Set testTile = MapManager.GetNextTile(CurrentTile, dRight)
        If testTile.IsTraversable Then
            DistToTarget = MapManager.TileDistance(this.TargetTile, testTile)
            If DistToTarget < minDist Then
                minDist = DistToTarget
                Set selectedTile = testTile
                selectedDirection = dRight
            End If
        End If
    End If
    
    Set this.CurrentTile = selectedTile
    this.Heading = selectedDirection
  
End Sub

Private Function GetIndexOfMin(numbers() As Long) As Long
    '//return the index of the smallest number in an array of numbers
    Dim i As Long
    Dim minVal As Long
    Dim indexOfMin As Long
    
    minVal = 2147483647
    
    For i = LBound(numbers) To UBound(numbers)
        If numbers(i) < minVal Then
            indexOfMin = i
            minVal = numbers(i)
        End If
    Next
    
    
End Function

Private Property Get IGamePiece_Speed() As Double
    IGamePiece_Speed = this.Speed
End Property
