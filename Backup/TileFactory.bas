Attribute VB_Name = "TileFactory"
'@Folder("PacManEngine.Maps")
Option Explicit
Private Const DEAULT_START_CELL As String = "C2"
Public Type Coord
    x As Integer
    y As Integer
End Type

Public Function NewTile(tileEncoding As String, xPos As Integer, yPos As Integer) As Tile
    Dim result As New Tile
    With result
        Select Case tileEncoding
            Case TileToken.WALL_TOKEN
                .Id = WALL_TOKEN
            Case TileToken.PELLET_TOKEN
            '// create a pellet
                .Id = PELLET_TOKEN
            Case TileToken.SUPER_PELLET_TOEKN
            '// create a SuperPellet
                .Id = SUPER_PELLET_TOEKN
            Case TileToken.OPEN_PATH
                .Id = OPEN_PATH
            Case TileToken.DOOR
                .Id = DOOR
        End Select
        .IsTraversable = tileEncoding <> TileToken.WALL_TOKEN And tileEncoding <> TileToken.DOOR
        .SetAddress xPos, yPos
    End With
    Set NewTile = result
End Function

Sub ClaimBoard(Optional startCell As Range)
    If startCell Is Nothing Then
        Set startCell = Range(DEAULT_START_CELL)
    Else
        Set startCell = startCell(1, 1)
    End If
    
    Dim fullBoard As Range
    Dim Column As Range
    Dim row As Range
    
    Set fullBoard = startCell.Resize(61, 55)
    
    For Each Column In fullBoard.Columns
        Column.ColumnWidth = 0.83
    Next
    
    For Each row In fullBoard.Rows
        row.RowHeight = 7.5
    Next
End Sub
