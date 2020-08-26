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
Private cycleCounter As Integer
Private Const MaxCycles As Integer = 5000
Private IsGameOver As Boolean

Public WithEvents UIAdapter As GameUIAdapter
Attribute UIAdapter.VB_VarHelpID = -1
Private UI As IGameUICommands

Public Pacman As PacmanController
Public Ghosts As Ghost '//collection of ghosts

Public Enum Direction
    dNone
    dUp
    dDown
    dLeft
    dRight
End Enum


Public Property Get Maze() As Tile()
    Maze = mMaze
End Property
Public Property Let Maze(value() As Tile)
    mMaze = value
    MapManager.Maze = mMaze
End Property

Private Sub BuildGameBoard()
    UI.DrawMap Me.Maze
End Sub

Private Sub Progress()
    
    
    If cycleCounter = Round(Pacman.Speed * MaxCycles, 0) Then
        Pacman.Move
        Debug.Print Pacman.CurrentTile.y & ", " & Pacman.CurrentTile.x & ": " & Pacman.CurrentTile.Id
    End If
    
    Debug.Assert Not Pacman.IsDead
    'If Pacman.IsDead Then IsGameOver = True
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
    
    Set Me.Pacman = New PacmanController
    Set Me.Pacman.CurrentTile = MapManager.GetMazeTile(46, 30)
    Me.Pacman.Heading = dLeft
    
    cycleCounter = 0
    
    Do While Not IsGameOver
        DoEvents
        If cycleCounter = MaxCycles Then cycleCounter = 0
        cycleCounter = cycleCounter + 1
        Progress
        DoEvents
        
    Loop
    
End Sub


'//User Input support
Private Sub UIAdapter_DownKeyPress()
    Me.Pacman.Heading = dDown
End Sub


Private Sub UIAdapter_GameStarted()
    '// initialize vars
        '//scoreboard
        '//
    '//play intro
    
    
End Sub

Private Sub UIAdapter_UpKeyPressed()
    Me.Pacman.Heading = dUp
End Sub
