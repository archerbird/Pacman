VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetViewWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ViewImplementations.ExcelWorksheet"
'//UI implemented as an Excel Worksheet
Option Explicit
Implements IViewCommands

Private Const MAP_START_ADDRESS As String = "$D$3"
Private MapRange As Range
Private WithEvents innerWs As Worksheet
Attribute innerWs.VB_VarHelpID = -1
Private mAdapter As IViewEvents
Private dPad As Range
Private ShapeWrappers As Dictionary

Private YIndexOffset As Long
Private XIndexOffset As Long

Public Enum KeyCode
    LeftArrow = 37
    RightArrow = 39
    UpArrow = 38
    DownArrow = 40
End Enum

Public Event Activated()
Public Event Deactivated()
Public Event DirectionalKeyPressed(vbKey As KeyCode)


Public Sub Init(xlWs As Worksheet)
    Dim s As Shape
    
    For Each s In xlWs.Shapes
        s.Delete
    Next
    
    xlWs.Activate
    xlWs.Range("AE65").Select
    Set innerWs = xlWs
    Set dPad = xlWs.Range("AE65")
End Sub

Private Sub Class_Terminate()
    Set mAdapter = Nothing
    Set innerWs = Nothing
    Set dPad = Nothing
     Debug.Print TypeName(Me) & " terminating..."
End Sub



Private Function TileToRange(mapTile As Tile) As Range
    Set TileToRange = MapRange.Cells(mapTile.y + YIndexOffset, mapTile.x + XIndexOffset)
End Function




'// Support for IGameUICommands
Private Sub IViewCommands_CreateGhost(character As Ghost)
    '// Create a corrosponding ViewModel Ghost
    Dim newGhostShape As New GhostStyler
    newGhostShape.Init innerWs, character.Color

    '// Add him to the drawing collection
    ShapeWrappers.Add character.Name, newGhostShape
    
End Sub

Private Sub IViewCommands_CreatePacman(character As PacmanController)
    '// Create a corrosponding ViewModel Pacman
    Dim newPacmanShape As New PacmanStyler
    newPacmanShape.Init innerWs
    
    '// Add him to the drawing collection
    ShapeWrappers.Add character.Name, newPacmanShape
    
End Sub

Private Sub IViewCommands_CreateMap(map() As Tile)

    YIndexOffset = 1 - LBound(map, 1)
    XIndexOffset = 1 - LBound(map, 2)
    
    Set MapRange = innerWs.Range(MAP_START_ADDRESS).Resize(UBound(map, 1) + YIndexOffset, UBound(map, 2) + XIndexOffset)
End Sub

Private Sub IViewCommands_UpdateComponent(characters As GamePieceCollection)
    Dim character As IGamePiece
    Dim characterShape As IDrawable
    
    For Each character In characters
        '// use the id from each character to get the corresponding ShapeWrapper
        Set characterShape = ShapeWrappers.Item(character.Id)
        characterShape.Redraw character.CurrentHeading, TileToRange(character.CurrentTile)
        
    Next
End Sub


Private Property Set IViewCommands_Events(ByVal RHS As IViewEvents)
    Set mAdapter = RHS
End Property

Private Property Get IViewCommands_Events() As IViewEvents
    Set IViewCommands_Events = mAdapter
End Property

'//Adds/Removes Keyboard Listening when the sheet does active/inactive
Private Sub innerWs_Activate()
    RaiseEvent Activated
End Sub

Private Sub innerWs_Deactivate()
    RaiseEvent Deactivated
End Sub

Private Sub innerWs_SelectionChange(ByVal Target As Range)
    If dPad.Offset(-1, 0).Address = Target.Address Then
        mAdapter.OnDirectionalKeyPress UpArrow
    ElseIf dPad.Offset(1, 0).Address = Target.Address Then
        mAdapter.OnDirectionalKeyPress (DownArrow)
    ElseIf dPad.Offset(0, -1).Address = Target.Address Then
        mAdapter.OnDirectionalKeyPress (LeftArrow)
    ElseIf dPad.Offset(0, 1).Address = Target.Address Then
        mAdapter.OnDirectionalKeyPress (RightArrow)
    End If
    
    Application.EnableEvents = False
    dPad.Select
    Application.EnableEvents = True
End Sub


