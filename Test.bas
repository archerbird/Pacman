Attribute VB_Name = "Test"
'@Folder("VBAProject")
Option Explicit

Sub test_shape_pos()
    Dim pm As Shape
    Dim cntr As Range
    
    Set pm = Sheet6.Shapes("PacMan")
    Set cntr = Range("C4")
    
    pm.Left = cntr.Left + ((cntr.Width - pm.Width) / 2)
    pm.Top = cntr.Top + ((cntr.height - pm.height) / 2)
End Sub

Sub LinkingInstructionStacking()
    
    Dim dict As Dictionary
    Dim fullBoard As Range
    Dim cell As Range
    
    Set fullBoard = Range("B2:BH59") 'Range("$B$2:$BB$60")
    Set dict = New Dictionary
    Set cell = fullBoard(1, 1)

    If Not dict.Exists(cell.Address) Then
        Dim gc As New GridNode
        Set gc.HostRange = cell
        dict.Add cell.Address, gc
        LinkCells gc, fullBoard, dict
    End If
    
    gc.Decompose
End Sub

Sub PickingUpTile()
    Dim testSubject As Tile
    Dim gb As GameGrid
    Set gb = New GameGrid
    Set gb.HostRange = ActiveSheet.Range("$B$2:$BB$60")
    Set TileFactory.GameBoard = gb
    
    Set testSubject = TileFactory.GetTile(Selection)
End Sub


