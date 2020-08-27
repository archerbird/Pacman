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
Private cycleCounter As Double
Private IsGameOver As Boolean

Public WithEvents UIAdapter As ViewAdapter
Attribute UIAdapter.VB_VarHelpID = -1
Private UI As IViewCommands

Public pacman As PacmanController
Public blinky As Ghost
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
    
    'Debug.Assert cycleCounter <> 2500
    If cycleCounter Mod Round(5 / pacman.Speed, 0) = 0 Then
        pacman.Move
        'Debug.Print "Frame: " & cycleCounter & " PM pos: " & pacman.CurrentTile.y & ", " & pacman.CurrentTile.x & ": " & pacman.CurrentTile.Id
    End If
    
    If cycleCounter Mod Round(5 / blinky.Speed, 0) = 0 Then
        blinky.Move
    End If
    
    'Debug.Assert Not Pacman.IsDead
    'If Pacman.IsDead Then IsGameOver = True
    UI.DrawGhost blinky
    UI.DrawPacman pacman
End Sub

Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " terminating..."
    UIAdapter.Deconstruct
    Set UIAdapter = Nothing
    Set UI = Nothing
End Sub

Public Sub StartGame(userInput As ViewAdapter)
    Set UI = userInput
    Set UIAdapter = userInput
    
    '// set up pacman
    Set Me.pacman = New PacmanController
    Set Me.pacman.CurrentTile = MapManager.GetMazeTile(46, 30)
    
    '// set up ghosts
    Set Me.blinky = New Ghost
    Set Me.blinky.CurrentTile = MapManager.GetMazeTile(22, 30)
    Set Me.blinky.ActiveBehavior = New ScatterBehavior
    
    
    UI.DrawMap mMaze
    
    cycleCounter = 0
    
    Do While Not IsGameOver
    
        DoEvents
        'If cycleCounter = MaxCycles Then cycleCounter = 0
        cycleCounter = cycleCounter + 1
        Progress
        'DoEvents
    Loop
    
End Sub


Private Sub UIAdapter_DirectionalKeyPressed(vbKeyCode As KeyCode)
    Select Case vbKeyCode
        Case KeyCode.UpArrow
            pacman.Heading = dUp
        Case KeyCode.DownArrow
            pacman.Heading = dDown
        Case KeyCode.LeftArrow
            pacman.Heading = dLeft
        Case KeyCode.RightArrow
            pacman.Heading = dRight
    End Select
    
End Sub

'//User Input support


Private Sub UIAdapter_GameStarted()
    '// initialize vars
        '//scoreboard
        '//
    '//play intro
    
    
End Sub

