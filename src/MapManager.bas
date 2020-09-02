Attribute VB_Name = "MapManager"
'@Folder "PacmanGame.Model.Maps"
Option Explicit

Private mMaze() As Tile


Public Property Get Maze() As Tile()
    Maze = mMaze
End Property
Public Property Let Maze(value() As Tile)
    mMaze = value
End Property

Public Property Get RowCount() As Long
    RowCount = (UBound(mMaze, 1) - LBound(mMaze, 1))
End Property

Public Property Get ColCount() As Long
    ColCount = (UBound(mMaze, 2) - LBound(mMaze, 2))
End Property
Public Function GetMazeTile(row As Integer, col As Integer) As Tile
    Set GetMazeTile = mMaze(SupressRow(row), SupressCol(col))
End Function

Public Function GetNextTile(CurrentTile As Tile, Heading As Direction, Optional lookAhead As Integer = 1) As Tile
        Select Case Heading
        Case Direction.dDown
            If CurrentTile.y = UBound(mMaze, 1) Then
            '//wrap around
                Set GetNextTile = mMaze(LBound(mMaze, 1) + lookAhead - 1, CurrentTile.x)
            Else
                Set GetNextTile = mMaze(SupressRow(CurrentTile.y + lookAhead), CurrentTile.x)
            End If
            
        Case Direction.dLeft
            If CurrentTile.x = LBound(mMaze, 2) Then
            '//wrap around
                Set GetNextTile = mMaze(CurrentTile.y, UBound(mMaze, 2) - lookAhead + 1)
            Else
                Set GetNextTile = mMaze(CurrentTile.y, SupressCol(CurrentTile.x - lookAhead))
            End If
            
        Case Direction.dRight
            If CurrentTile.x = UBound(mMaze, 2) Then
            '//wrap around
                Set GetNextTile = mMaze(CurrentTile.y, LBound(mMaze, 2) + lookAhead - 1)
            Else
                Set GetNextTile = mMaze(CurrentTile.y, SupressCol(CurrentTile.x + lookAhead))
            End If
            
        Case Direction.dUp
            If CurrentTile.y = LBound(mMaze, 1) Then
            '//wrap around
                Set GetNextTile = mMaze(UBound(mMaze, 1) - (lookAhead + 1), CurrentTile.x)
            Else
                Set GetNextTile = mMaze(SupressRow(CurrentTile.y - lookAhead), CurrentTile.x)
            End If
    End Select
End Function

Public Function TileDistance(targetedTile As Tile, optionTile As Tile) As Long
    '// a^2 +b^2 = c^2

    TileDistance = Sqr((targetedTile.y - optionTile.y) ^ 2 + (targetedTile.x - optionTile.x) ^ 2)

End Function

Private Function SupressRow(row As Integer) As Integer
    If row <= UBound(mMaze, 1) And row >= LBound(mMaze, 1) Then
        SupressRow = row
    ElseIf row > UBound(mMaze, 1) Then
        SupressRow = UBound(mMaze, 1)
    ElseIf row < LBound(mMaze, 1) Then
        SupressRow = LBound(mMaze, 1)
    End If
End Function

Private Function SupressCol(col As Integer) As Integer
    If col <= UBound(mMaze, 2) And col >= LBound(mMaze, 2) Then
        SupressCol = col
    ElseIf col > UBound(mMaze, 2) Then
        SupressCol = UBound(mMaze, 2)
    ElseIf col < LBound(mMaze, 2) Then
        SupressCol = LBound(mMaze, 2)
    End If
End Function


Private Function CyclicRow(row As Integer) As Integer
    
    If row <= UBound(mMaze, 1) And row >= LBound(mMaze, 1) Then
        CyclicRow = row
    ElseIf row > UBound(mMaze, 1) Then
        CyclicRow = CyclicRow(row - RowCount)
    ElseIf row < LBound(mMaze, 1) Then
        CyclicRow = CyclicRow(row + RowCount)
    End If
    
End Function

Private Function CyclicCol(col As Integer) As Integer
    If col <= UBound(mMaze, 2) And col >= LBound(mMaze, 2) Then
        CyclicCol = col
    ElseIf col > UBound(mMaze, 2) Then
        CyclicCol = CyclicCol(col - ColCount)
    ElseIf col < LBound(mMaze, 2) Then
        CyclicCol = CyclicCol(col + ColCount)
    End If
End Function

Public Function LoadMapFromFile() As Tile()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Maps\defaultMap.pmap"
    
    LoadMapFromFile = TransformToMap(ReadText(filePath))
    
End Function

Private Function TransformToMap(inputString As String) As Tile()
    Dim rowArr As Variant
    Dim element As Variant
    Dim subElement As Variant
    Dim result() As Tile
    Dim subArr() As String
    Dim RowCount As Integer
    Dim ColCount As Integer
    Dim j As Integer
    Dim i As Integer
    
    rowArr = Split(inputString, ";")
    RowCount = UBound(rowArr) - LBound(rowArr) + 1
    ColCount = UBound(Split(rowArr(LBound(rowArr)), ",")) - LBound(Split(rowArr(LBound(rowArr)), ",")) + 1
    
    ReDim result(1 To RowCount, 1 To ColCount)
    
    For Each element In rowArr
        j = j + 1
        i = 0
        subArr = Split(element, ",")
        For Each subElement In subArr
            i = i + 1
            Set result(j, i) = TileFactory.NewTile(CStr(subElement), i, j)
        Next
    Next
    
    TransformToMap = result
End Function



Private Function ReadText(fileName As String) As String
    Dim textLine As String
    
    Open fileName For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textLine
        ReadText = ReadText & textLine
    Loop
    
    Close #1
End Function

