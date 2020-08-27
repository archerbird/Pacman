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

Dim tempMoveCount As Long
'Private mMaze() As Tile
'Private CycleCounter As Double
Private IsGameOver As Boolean
Private Const n As Double = 10

Public WithEvents UIAdapter As ViewAdapter
Attribute UIAdapter.VB_VarHelpID = -1
Private UI As IViewCommands

Public pacman As PacmanController
Public blinky As Ghost
Private mGhosts As Collection '//collection of ghosts

Public Enum Direction
    dNone = 0
    dUp = -1
    dDown = 1
    dLeft = -2
    dRight = 2
End Enum

Private Type TGameController
    Maze() As Tile
    CycleCounter As Long
    Ghosts As Collection
End Type

Private this As TGameController

Public Sub AddGhost(pmGhost As Ghost)
    this.Ghosts.Add pmGhost
End Sub

Public Property Get Maze() As Tile()
    Maze = this.Maze 'mMaze
End Property
Public Property Let Maze(value() As Tile)
    this.Maze = value
    MapManager.Maze = this.Maze
End Property

Private Sub BuildGameBoard()
    UI.DrawMap Me.Maze
End Sub

Private Sub Progress()
    Sleep 0.05
    
    If pacman.CycleRemainder >= n Then
        pacman.CycleRemainder = pacman.CycleRemainder Mod n
        pacman.Move
        tempMoveCount = tempMoveCount + 1
                
    
    Else
    'Debug.Assert cycleCounter <> 2500
        If this.CycleCounter Mod Round(n / (n * (1 - pacman.Speed)), 0) <> 0 Then
        
           pacman.CycleRemainder = pacman.CycleRemainder + n Mod (n * (1 - pacman.Speed))
           pacman.Move
           
           tempMoveCount = tempMoveCount + 1
               
        End If
        
        If Round(n / (n * (1 - pacman.Speed)), 0) = 1 Then
           pacman.CycleRemainder = pacman.CycleRemainder + n Mod (n * (1 - pacman.Speed))
    
        End If
        'Debug.Print "Frame: " & cycleCounter & " PM pos: " & pacman.CurrentTile.y & ", " & pacman.CurrentTile.x & ": " & pacman.CurrentTile.Id


    End If
    
                Debug.Print "Frame: " & this.CycleCounter
                Debug.Print "Moves: " & tempMoveCount
                Debug.Print "CycleRemainder: " & pacman.CycleRemainder
                Debug.Print "-----------------------"
                
    If this.CycleCounter Mod Round(n / (n * (1 - blinky.Speed)), 0) <> 0 Then
        blinky.Move
    End If
    
    'Debug.Assert Not cycleCounter = 100
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
    'Set Me.blinky.ActiveBehavior = New ScatterBehavior
    Set Me.blinky.ActiveBehavior = ChasePacmanBehavior.Create(Me.pacman)
    
    
    UI.DrawMap this.Maze
    
    this.CycleCounter = 0
    
    Do While Not IsGameOver
    
        'DoEvents
        'If cycleCounter = MaxCycles Then cycleCounter = 0
        this.CycleCounter = this.CycleCounter + 1
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

Private Sub Sleep(sec As Double)
    Dim t As Double
    
    t = Timer
    
    Do
    DoEvents
    Loop Until Timer > t + sec
End Sub


