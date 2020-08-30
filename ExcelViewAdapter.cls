VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelViewAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''@Folder "ViewImplementations.ExcelWorksheet"
'Option Explicit
'Implements IViewCommands
'Implements IViewEvents
'
''//two ways to access the same object
'Private WithEvents view As WorksheetViewWrapper
'Private viewCommand As IViewCommands
'Private mAdapter As IViewEvents
'
''Private WithEvents keyListener As KeyPressDispatcher
'
'Public Sub Init(sheetUI As WorksheetViewWrapper)
'    Set view = sheetUI
'    Set viewCommand = sheetUI
'End Sub
'
'Private Sub Class_Terminate()
'    Set mAdapter = Nothing
'    Set view = Nothing
'    Set viewCommand = Nothing
'    Debug.Print TypeName(Me) & " terminating..."
'End Sub
'
'
'
'Private Sub IViewCommands_CreateGhost(character As Ghost)
'
'End Sub
'
'Private Sub IViewCommands_CreateMap(map() As Tile)
'
'End Sub
'
'Private Sub IViewCommands_CreatePacman(character As PacmanController)
'
'End Sub
'
''//Delegate commands down to the WorksheetUIWrapper
'
'Private Sub IViewCommands_DrawMap(map() As Tile)
'    viewCommand.DrawMap map
'End Sub
'
'Private Sub IViewCommands_CreateGhostViewModel(character As Ghost)
'    viewCommand.CreateGhostViewModel character
'End Sub
'
'Private Sub IViewCommands_CreatePacmanViewModel(character As PacmanController)
'    viewCommand.CreatePacmanViewModel character
'End Sub
'Private Property Set IViewCommands_Events(ByVal RHS As ViewAdapter)
'    Set mAdapter = RHS
'End Property
'
'Private Property Get IViewCommands_Events() As ViewAdapter
'    Set IViewCommands_Events = mAdapter
'End Property
'
'Private Sub IViewCommands_RedrawCharacters()
'    viewCommand.RedrawCharacters
'End Sub
'
'Private Sub IViewCommands_RedrawTile(updatedTile As Tile)
'    viewCommand.RedrawTile updatedTile
'End Sub
'
'Private Sub IViewCommands_UpdateComponent(characters As GamePieceCollection)
'
'End Sub
'
''// Events
'Private Sub IViewEvents_OnDirectionalKeyPress(vbKey As KeyCode)
'    mAdapter.OnDirectionalKeyPress vbKey
'End Sub
'
'Private Sub IViewEvents_OnGameStarted()
'    mAdapter.OnGameStarted
'End Sub
'
'
''// Tells the keyListener when to start/stop raising events
'Private Sub view_Activated()
'
'End Sub
'
'Private Sub view_DirectionalKeyPressed(vbKey As KeyCode)
'    mAdapter.OnDirectionalKeyPress vbKey
'End Sub
