Attribute VB_Name = "TileFactory"
'@Folder("PacManEngine.Maps")
Option Explicit
Private Const DEAULT_START_CELL As String = "C2"

Public Function NewTile(tileEncoding As String) As Tile
    Dim result As New Tile
    With result
        Select Case tileEncoding
            Case TileToken.WALL_TOKEN
                .Id = "Wall"
            Case TileToken.PELLET_TOKEN
            '// create a pellet
                .Id = "Pellet"
            Case TileToken.SUPER_PELLET_TOEKN
            '// create a SuperPellet
                .Id = "SuperPellet"
            Case TileToken.OPEN_PATH
                .Id = "Path"
        End Select
        .IsTraversable = tileEncoding <> TileToken.WALL_TOKEN
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
    Dim Row As Range
    
    Set fullBoard = startCell.Resize(61, 55)
    
    For Each Column In fullBoard.Columns
        Column.ColumnWidth = 0.83
    Next
    
    For Each Row In fullBoard.Rows
        Row.RowHeight = 7.5
    Next
End Sub
