Attribute VB_Name = "Test"
'@Folder("VBAProject")
Option Explicit

Sub Test_OuterEdgeDrawing()
    Dim gCell As GridCell
    Dim gCell2 As GridCell
    Dim rng As Range
    
    Set rng = Selection.Offset(1, 0)
    Set gCell = GridCellFactory.OuterWallGrid(Selection)
    Set gCell2 = GridCellFactory.OuterWallGrid(rng)
    
    Set gCell.LowerPathNode = gCell2
    Drawing.Render gCell
    Drawing.Render gCell2
End Sub

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
        Dim gc As New GridCell
        Set gc.HostRange = cell
        dict.Add cell.Address, gc
        LinkCells gc, fullBoard, dict
    End If
    
    gc.Decompose
End Sub

Sub LinkCells(gc As GridCell, fb As Range, dict As Dictionary)

    
    '//Link Top
    LinkTop gc, fb, dict
   
    '//Link Right

    LinkRight gc, fb, dict

    
''''    '//Link Bottom
''''    isWorksheetLimit = gc.HostRange.row = gc.HostRange.Parent.Rows.Count
''''
''''    If Not isWorksheetLimit Then
''''        isBoardLimit = Application.Intersect(gc.HostRange.Offset(1, 0), fb) Is Nothing
''''    End If
''''
''''    If isWorksheetLimit Or isBoardLimit Then
''''        '// need to wrap around to bottom
''''        Set target = fb(1, gc.HostRange.column - fb.column + 1) '// wraps to the first row, same column
''''    Else
''''        '// simply take the cell just below
''''        Set target = gc.HostRange.Offset(1, 0)
''''    End If
''''
''''    Dim ngc As GridCell
''''    If Not dict.Exists(target.Address) Then
''''        '// no GridCell has been created for this range yet
''''        Set ngc = New GridCell
''''        With ngc
''''            Set .HostRange = target
''''            dict.Add target.Address, ngc
''''            LinkCells gc, fb, dict
''''        End With
''''
''''    Else
''''        '// a grid cell has been make already
''''        Set ngc = dict(target.Address)
''''    End If
''''
''''    If ngc.LowerPathNode Is Nothing Then
''''        '//if no node has been set here, we are safe to add it to our node
''''        Set ngc.UpperPathNode = gc
''''        Set gc.LowerPathNode = ngc
''''    End If
''''        '// other wise the upward recurse stops here
End Sub

Sub LinkTop(gc As GridCell, fb As Range, dict As Dictionary)
    Dim target As Range
    Dim isWorksheetLimit As Boolean
    Dim isBoardLimit As Boolean
    Dim ngc As GridCell
    
'//Link Top
    isWorksheetLimit = gc.HostRange.row = 1
    
    If Not isWorksheetLimit Then
        isBoardLimit = Application.Intersect(gc.HostRange.Offset(-1, 0), fb) Is Nothing
    End If
    
    If isWorksheetLimit Or isBoardLimit Then
        '// need to wrap around to bottom
        Set target = fb(fb.Rows.Count, gc.HostRange.column - fb.column + 1) '// wraps to the last row, same column
    Else
        '// simply take the cell just above
        Set target = gc.HostRange.Offset(-1, 0)
    End If

    If Not dict.Exists(target.Address) Then
        '// no GridCell has been created for this range yet
        Set ngc = New GridCell
        With ngc
            Set .HostRange = target
            dict.Add target.Address, ngc
        End With
        
    Else
        '// a grid cell has been make already
        Set ngc = dict(target.Address)
    End If
    
    If ngc.LowerPathNode Is Nothing Then
        '//if no node has been set here, we are safe to add it to our node
        Set ngc.LowerPathNode = gc
        Set gc.UpperPathNode = ngc
        LinkTop ngc, fb, dict
    End If
        '// other wise the upward recurse stops here
End Sub

Sub LinkRight(gc As GridCell, fb As Range, dict As Dictionary)
    Dim target As Range
    Dim isWorksheetLimit As Boolean
    Dim isBoardLimit As Boolean
    Dim ngc As GridCell
    
'//Link Right
    isWorksheetLimit = gc.HostRange.column = gc.HostRange.Parent.Columns.Count
    
    If Not isWorksheetLimit Then
        isBoardLimit = Application.Intersect(gc.HostRange.Offset(0, 1), fb) Is Nothing
    End If
    
    If isWorksheetLimit Or isBoardLimit Then
        '// need to wrap around to right
        Set target = fb(gc.HostRange.row - fb.row + 1, 1) '//wraps to the same row, first column
    Else
        '// simply take the cell to the right
        Set target = gc.HostRange.Offset(0, 1)
    End If
    
    
    If Not dict.Exists(target.Address) Then
        '// no GridCell has been created for this range yet
        Set ngc = New GridCell
        With ngc
            Set .HostRange = target
        End With
        
    Else
        '// a grid cell has been make already
        Set ngc = dict(target.Address)
    End If
    
    If ngc.LeftPathNode Is Nothing Then
        '//if no node has been set here, we are safe to add it to our node
        Set ngc.LeftPathNode = gc
        Set gc.RightPathNode = ngc
        LinkRight ngc, fb, dict
    End If
End Sub
