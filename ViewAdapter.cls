VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ViewAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.View"
'//"Abstract" class
Option Explicit
Implements IViewEvents
Implements IViewCommands


Private mCommandsUI As IViewCommands

Public Event DirectionalKeyPressed(vbKeyCode As KeyCode)
Public Event GameStarted()
Public Event GamePaused()
Public Event Quit()



Public Sub Init(viewUI As IViewCommands)
    Set mCommandsUI = viewUI
    Set viewUI.Events = Me
End Sub

Public Sub Deconstruct()
    Set mCommandsUI.Events = Nothing
End Sub

Private Sub Class_Terminate()
    Set mCommandsUI = Nothing
    Debug.Print TypeName(Me) & " terminating..."
End Sub

Public Function AsCommandSender() As IViewCommands
    Set AsCommandSender = Me
End Function

'//IGameUICommands Support
Private Sub IViewCommands_CreateGhost(character As Ghost)
    mCommandsUI.CreateGhost character
End Sub

Private Sub IViewCommands_CreatePacman(character As PacmanController)
    mCommandsUI.CreatePacman character
End Sub

Private Sub IViewCommands_CreateMap(map() As Tile)
    mCommandsUI.CreateMap map
End Sub

Private Sub IViewCommands_UpdateComponent(characters As Collection)
    mCommandsUI.UpdateComponent characters
End Sub

'//IGameUIEvents Support
Private Sub IViewEvents_OnDirectionalKeyPress(vbKey As KeyCode)
    RaiseEvent DirectionalKeyPressed(vbKey)
End Sub

Private Property Set IViewCommands_Events(ByVal RHS As IViewEvents)
    '//never let this actually get set
End Property

Private Property Get IViewCommands_Events() As IViewEvents
    Set IViewCommands_Events = Me
End Property

Private Sub IViewEvents_OnGamePaused()
    RaiseEvent GamePaused
End Sub

Private Sub IViewEvents_OnGameStarted()
    RaiseEvent GameStarted
End Sub

Private Sub IViewEvents_OnQuit()
    RaiseEvent Quit
End Sub
