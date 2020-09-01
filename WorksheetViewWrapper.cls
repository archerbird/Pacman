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
Implements IGameEventsHandler

Private Const MAP_START_ADDRESS As String = "$D$3"

Private Type TWorksheetViewWrapper
    MapRange As Range
    dPad As Range
    Adapter As IViewEventsHandler
    ShapeWrappers As Dictionary
    YIndexOffset As Long
    XIndexOffset As Long
End Type

Private WithEvents innerWs As Worksheet
Attribute innerWs.VB_VarHelpID = -1
Private this As TWorksheetViewWrapper

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
    Set this.dPad = xlWs.Range("AE65")
End Sub


Private Sub Class_Initialize()
    Set this.ShapeWrappers = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set this.Adapter = Nothing
    Set innerWs = Nothing
    Set this.dPad = Nothing
     Debug.Print TypeName(Me) & " terminating..."
End Sub


'// Support for IGameEventsHandler
Private Sub IGameEventsHandler_CreateGhost(character As GhostModel)
    '// Create a corrosponding ViewModel Ghost
    Dim newGhostShape As New GhostStyler
    newGhostShape.Init innerWs, character.Color

    '// Add him to the drawing collection
    this.ShapeWrappers.Add character.Name, newGhostShape
    
End Sub

Private Sub IGameEventsHandler_CreatePacman(character As PacmanModel)
    '// Create a corrosponding ViewModel Pacman
    Dim newPacmanShape As New PacmanStyler
    newPacmanShape.Init innerWs
    
    '// Add him to the drawing collection
    this.ShapeWrappers.Add character.Name, newPacmanShape
    
End Sub

Private Sub IGameEventsHandler_CreateMap(map() As Tile)

    this.YIndexOffset = 1 - LBound(map, 1)
    this.XIndexOffset = 1 - LBound(map, 2)
    
    Set this.MapRange = innerWs.Range(MAP_START_ADDRESS).Resize(UBound(map, 1) + this.YIndexOffset, UBound(map, 2) + this.XIndexOffset)
End Sub

Private Sub IGameEventsHandler_UpdateComponents(characters As Collection)
    Dim character As IGamePiece
    Dim characterShape As IDrawable
    Dim i As Integer
    
    For Each character In characters
        '// use the id from each character to get the corresponding ShapeWrapper
        Set characterShape = this.ShapeWrappers.Item(character.Id)
        characterShape.Redraw character.CurrentHeading, TileToRange(character.CurrentTile)
        
    Next
End Sub

Private Property Set IGameEventsHandler_Events(ByVal RHS As IViewEventsHandler)
    Set this.Adapter = RHS
End Property

Private Property Get IGameEventsHandler_Events() As IViewEventsHandler
    Set IGameEventsHandler_Events = this.Adapter
End Property



'// Events from the worksheet that we will translate into view events
Private Sub innerWs_Activate()
    '// maybe pause the game?
End Sub

Private Sub innerWs_Deactivate()
    '// maybe we need a resume game event?
End Sub

Private Sub innerWs_SelectionChange(ByVal Target As Range)
    If this.dPad.Offset(-1, 0).Address = Target.Address Then
        this.Adapter.OnDirectionalKeyPress UpArrow
    ElseIf this.dPad.Offset(1, 0).Address = Target.Address Then
        this.Adapter.OnDirectionalKeyPress (DownArrow)
    ElseIf this.dPad.Offset(0, -1).Address = Target.Address Then
        this.Adapter.OnDirectionalKeyPress (LeftArrow)
    ElseIf this.dPad.Offset(0, 1).Address = Target.Address Then
        this.Adapter.OnDirectionalKeyPress (RightArrow)
    End If
    
    Application.EnableEvents = False
    this.dPad.Select
    Application.EnableEvents = True
End Sub


'// Private helpers
Private Function TileToRange(mapTile As Tile) As Range
    Set TileToRange = this.MapRange.Cells(mapTile.y + this.YIndexOffset, mapTile.x + this.XIndexOffset)
End Function
