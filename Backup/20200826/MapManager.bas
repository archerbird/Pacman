Attribute VB_Name = "MapManager"
'@Folder "PacManEngine.Model.Maps"
Option Explicit

Private mMaze() As Tile


Public Property Get Maze() As Tile()
    Maze = mMaze
End Property
Public Property Let Maze(value() As Tile)
    mMaze = value
End Property
Public Function GetMazeTile(row As Integer, col As Integer) As Tile
    Set GetMazeTile = mMaze(row, col)
End Function

Public Function GetNextTile(CurrentTile As Tile, Heading As Direction, Optional lookAhead As Integer = 1) As Tile
        Select Case Heading
        Case Direction.dDown
            If CurrentTile.y = UBound(mMaze, 1) Then
            '//wrap around
                Set GetNextTile = mMaze(LBound(mMaze, 1) + lookAhead - 1, CurrentTile.x)
            Else
                Set GetNextTile = mMaze(CurrentTile.y + lookAhead, CurrentTile.x)
            End If
            
        Case Direction.dLeft
            If CurrentTile.x = LBound(mMaze, 2) Then
            '//wrap around
                Set GetNextTile = mMaze(CurrentTile.y, UBound(mMaze, 2) - lookAhead + 1)
            Else
                Set GetNextTile = mMaze(CurrentTile.y, CurrentTile.x - lookAhead)
            End If
            
        Case Direction.dRight
            If CurrentTile.x = UBound(mMaze, 2) Then
            '//wrap around
                Set GetNextTile = mMaze(CurrentTile.y, LBound(mMaze, 2) + lookAhead - 1)
            Else
                Set GetNextTile = mMaze(CurrentTile.y, CurrentTile.x + lookAhead)
            End If
            
        Case Direction.dUp
            If CurrentTile.y = LBound(mMaze, 1) Then
            '//wrap around
                Set GetNextTile = mMaze(UBound(mMaze, 1) - (lookAhead + 1), CurrentTile.x)
            Else
                Set GetNextTile = mMaze(CurrentTile.y - lookAhead, CurrentTile.x)
            End If
    End Select
End Function

