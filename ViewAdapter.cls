VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacmanGame.View"
Option Explicit
Implements IViewEventsHandler
Implements IGameEventsHandler

Private Const mModuleName As String = "ViewAdapter"
Private viewUI As IGameEventsHandler

Public Event DirectionalKeyPressed(vbKeyCode As KeyCode)
Public Event GameStarted()
Public Event GamePaused()
Public Event Quit()


Public Sub Init(inViewUI As IGameEventsHandler)
    Set viewUI = inViewUI
    Set viewUI.Events = Me
End Sub

Public Sub Deconstruct()
    '// unhooks itself from the GameEventsHandler to prevent memory leakage
    Set viewUI.Events = Nothing
End Sub

Public Function AsCommandSender() As IGameEventsHandler
    '// allows access to the IGameEventsHandler methods
    Set AsCommandSender = Me
End Function

Private Sub Class_Terminate()
    Set viewUI = Nothing
    Debug.Print TypeName(Me) & " terminating..."
End Sub


'//IGameEventsHandler Support
Private Property Set IGameEventsHandler_Events(ByVal RHS As IViewEventsHandler)
    '//this isn't meant to be set from the outside for this class
End Property

Private Property Get IGameEventsHandler_Events() As IViewEventsHandler
    Set IGameEventsHandler_Events = Me
End Property

Private Sub IGameEventsHandler_CreateGhost(character As GhostModel)
    viewUI.CreateGhost character
End Sub

Private Sub IGameEventsHandler_CreatePacman(character As PacmanModel)
    viewUI.CreatePacman character
End Sub

Private Sub IGameEventsHandler_CreateMap(map() As Tile)
    viewUI.CreateMap map
End Sub

Private Sub IGameEventsHandler_UpdateComponents(characters As Collection)
    viewUI.UpdateComponents characters
End Sub


'//IViewEventsHandler Support
Private Sub IViewEventsHandler_OnDirectionalKeyPress(vbKey As KeyCode)
    RaiseEvent DirectionalKeyPressed(vbKey)
End Sub

Private Sub IViewEventsHandler_OnGamePaused()
    RaiseEvent GamePaused
End Sub

Private Sub IViewEventsHandler_OnGameStarted()
    RaiseEvent GameStarted
End Sub

Private Sub IViewEventsHandler_OnQuit()
    RaiseEvent Quit
End Sub
