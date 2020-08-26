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

'//IGameUICommands Support
Private Sub IViewCommands_DrawGhost(character As Ghost)
    mCommandsUI.DrawGhost character
End Sub

Private Sub IViewCommands_DrawMap(map() As Tile)
    mCommandsUI.DrawMap map
End Sub

Private Sub IViewCommands_DrawPacman(character As PacmanController)
    mCommandsUI.DrawPacman character
End Sub




Private Property Set IViewCommands_Events(ByVal RHS As ViewAdapter)
    '//never let this actually get set
End Property

Private Property Get IViewCommands_Events() As ViewAdapter
    Set IViewCommands_Events = Me
End Property

Private Sub IViewEvents_OnDirectionalKeyPress(vbKey As KeyCode)
    RaiseEvent DirectionalKeyPressed(vbKey)
End Sub

'//IGameUIEvents Support
'Private Sub IGameUIEvents_OnDirectionalKeyPress(vbKey As KeyCodes)
'    RaiseEvent DirectionalKeyPressed(vbKey)
'End Sub

Private Sub IViewEvents_OnGameStarted()
    RaiseEvent GameStarted
End Sub
