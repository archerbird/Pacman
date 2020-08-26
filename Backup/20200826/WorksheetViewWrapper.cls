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
Private mAdapter As ViewAdapter
Private dPad As Range
Private pacman As PacmanStyler
Private blinky As GhostStyler

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
    pacman.CreateShape innerWs
End Sub

Private Sub Class_Initialize()
    Set pacman = New PacmanStyler
    Set blinky = New GhostStyler
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
Private Sub IViewCommands_DrawGhost(character As Ghost)
    blinky.Redraw character.Heading, TileToRange(character.CurrentTile)
End Sub

Private Sub IViewCommands_DrawMap(map() As Tile)

    YIndexOffset = 1 - LBound(map, 1)
    XIndexOffset = 1 - LBound(map, 2)
    
    Set MapRange = innerWs.Range(MAP_START_ADDRESS).Resize(UBound(map, 1) + YIndexOffset, UBound(map, 2) + XIndexOffset)
End Sub

Private Sub IViewCommands_DrawPacman(character As PacmanController)
    pacman.Redraw character.Heading, TileToRange(character.CurrentTile)
End Sub




Private Property Set IViewCommands_Events(ByVal RHS As ViewAdapter)
    Set mAdapter = RHS
End Property

Private Property Get IViewCommands_Events() As ViewAdapter
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
        RaiseEvent DirectionalKeyPressed(UpArrow)
    ElseIf dPad.Offset(1, 0).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(DownArrow)
    ElseIf dPad.Offset(0, -1).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(LeftArrow)
    ElseIf dPad.Offset(0, 1).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(RightArrow)
    End If
    
    Application.EnableEvents = False
    dPad.Select
    Application.EnableEvents = True
End Sub
