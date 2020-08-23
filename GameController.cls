VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine")
'@Exposed
Option Explicit

Private mMaze() As Tile

Public WithEvents UIAdapter As GameUIAdapter
Attribute UIAdapter.VB_VarHelpID = -1
Private UI As IGameUICommands

Public Pacman As Pacman
Public Ghosts As Ghost '//collection of ghosts


Public Property Get Maze() As Tile()
    Maze = mMaze
End Property
Public Property Let Maze(value() As Tile)
    mMaze = value
End Property

Private Sub BuildGameBoard()
    UI.DrawMap mMaze
End Sub

Private Sub Progress()
    '//Move all characters
    
End Sub

Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " terminating..."
    UIAdapter.Deconstruct
    Set UIAdapter = Nothing
    Set UI = Nothing
End Sub

Public Sub StartGame(userInput As GameUIAdapter)
    Set UI = userInput
    Set UIAdapter = userInput
End Sub


'//User Input support
Private Sub UIAdapter_DownKeyPress()
MsgBox "Pressed Down"
End Sub

