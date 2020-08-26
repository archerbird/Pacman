VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelUIAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
Option Explicit
Implements IGameUICommands
Implements IGameUIEvents

'//two ways to access the same object
Private WithEvents sheetUI As WorksheetUIWrapper
Attribute sheetUI.VB_VarHelpID = -1
Private sheetUICommand As IGameUICommands
Private mAdapter As IGameUIEvents

'Private WithEvents keyListener As KeyPressDispatcher

Public Sub Init(sheetUI As WorksheetUIWrapper)
    Set sheetUI = sheetUI
    Set sheetUICommand = sheetUI
End Sub

Private Sub Class_Terminate()
    Set mAdapter = Nothing
    Set sheetUI = Nothing
    Set sheetUICommand = Nothing
    Debug.Print TypeName(Me) & " terminating..."
End Sub

'//Delegate commands down to the WorksheetUIWrapper
Private Sub IGameUICommands_DrawGhost()
    sheetUICommand.DrawGhost
End Sub

Private Sub IGameUICommands_DrawMap(map() As Tile)
    sheetUICommand.DrawMap map
End Sub

Private Sub IGameUICommands_DrawPacman()
    sheetUICommand.DrawPacman
End Sub


Private Property Set IGameUICommands_Events(ByVal RHS As GameUIAdapter)
    Set mAdapter = RHS
End Property

Private Property Get IGameUICommands_Events() As GameUIAdapter
    Set IGameUICommands_Events = mAdapter
End Property

'//Pass along event notifiers from the keyListener
Private Sub IGameUIEvents_OnDirectionalKeyPress(vbKey As KeyCode)
    mAdapter.OnDirectionalKeyPress vbKey
End Sub

Private Sub IGameUIEvents_OnGameStarted()
    mAdapter.OnGameStarted
End Sub

Private Sub keyListener_DirectionalKeyPressed(vbKey As KeyCode)
    IGameUIEvents_OnDirectionalKeyPress vbKey
End Sub

'// Tells the keyListener when to start/stop raising events
Private Sub sheetUI_Activated()

End Sub

