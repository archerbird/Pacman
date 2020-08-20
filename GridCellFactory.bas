Attribute VB_Name = "GridCellFactory"
'@Folder("PacManEngine.GridCells")
Option Explicit
Private Const DEAULT_START_CELL As String = "B2"

Public Function OuterWallGrid(host As Range) As GridCell
    Set OuterWallGrid = New GridCell
    With OuterWallGrid
        Set .HostRange = host(1, 1)
        Set .Pen = Drawing.OutsideBorderLine()
        .IsTraversable = False
    End With
End Function

Public Function InnerWallGrid(host As Range) As GridCell
    Set InnerWallGrid = New GridCell
    With InnerWallGrid
        Set .HostRange = host(1, 1)
        Set .Pen = Drawing.InsideBorderLine()
        .IsTraversable = False
    End With
End Function

Public Function PathGrid(host As Range) As GridCell
    Set PelletGrid = New GridCell
    With PelletGrid
        Set .HostRange = host
        .IsTraversable = True
    End With
End Function

Public Sub ConnectGridCells(gridCell1 As GridCell, gridCell2 As GridCell)
    Set gridCell1.LeftPathNode = gridCell2
    Set gridCell2.RightPathNode = gridCell1
End Sub

Sub ClaimBoard(Optional startCell As Range)
    If startCell Is Nothing Then
        Set startCell = Range(DEAULT_START_CELL)
    Else
        Set startCell = startCell(1, 1)
    End If
    
    Dim fullBoard As Range
    Dim column As Range
    Dim row As Range
    
    Set fullBoard = startCell.Resize(59, 53)
    
    For Each column In fullBoard.Columns
        column.ColumnWidth = 0.83
    Next
    
    For Each row In fullBoard.Rows
        row.RowHeight = 7.5
    Next
End Sub
