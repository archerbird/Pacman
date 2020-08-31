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
Private player As PacmanController

Public WithEvents UIAdapter As ViewAdapter
Attribute UIAdapter.VB_VarHelpID = -1

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
    Moveables As Collection
    Pacman As PacmanController
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
    UIAdapter.AsCommandSender.CreateMap Me.Maze
End Sub

Private Sub Progress()
    Sleep 0.05
    
    Dim character As IGamePiece
    
    
    For Each character In this.Moveables
    
        If character.CycleRemainder >= n Then
            character.CycleRemainder = character.CycleRemainder Mod n
            character.Move
            tempMoveCount = tempMoveCount + 1
                    
        
        Else
        'Debug.Assert cycleCounter <> 2500
            If this.CycleCounter Mod Round(n / (n * (1 - character.Speed)), 0) <> 0 Then
            
               character.CycleRemainder = character.CycleRemainder + n Mod (n * (1 - character.Speed))
               character.Move
               
               tempMoveCount = tempMoveCount + 1
                   
            End If
            
            If Round(n / (n * (1 - character.Speed)), 0) = 1 Then
               character.CycleRemainder = character.CycleRemainder + n Mod (n * (1 - character.Speed))
        
            End If
            'Debug.Print "Frame: " & cycleCounter & " PM pos: " & pacman.CurrentTile.y & ", " & pacman.CurrentTile.x & ": " & pacman.CurrentTile.Id
    
    
        End If
    Next
    
'                Debug.Print "Frame: " & this.CycleCounter
'                Debug.Print "Moves: " & tempMoveCount
'                Debug.Print "CycleRemainder: " & Pacman.CycleRemainder
'                Debug.Print "-----------------------"
    
    'Debug.Assert Not cycleCounter = 100
    'If Pacman.IsDead Then IsGameOver = True
    'UI.DrawGhost this.Ghosts
    UIAdapter.AsCommandSender.UpdateComponent this.Moveables
End Sub

Private Sub Class_Initialize()
    Set this.Moveables = New Collection
End Sub

Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " terminating..."
    Set this.Moveables = Nothing
    
    UIAdapter.Deconstruct
    
    Erase this.Maze
    Erase MapManager.Maze

    Set UIAdapter = Nothing
End Sub

Public Sub StartGame(userInput As ViewAdapter)
    Dim blinky As Ghost
    Dim inky As Ghost

    Set UIAdapter = userInput
    
    '// set up pacman
    Set player = New PacmanController
    Set player.CurrentTile = MapManager.GetMazeTile(46, 30)
    this.Moveables.Add player
    UIAdapter.AsCommandSender.CreatePacman player
    
    '// set up ghosts
    Set blinky = New Ghost
    blinky.Color = vbRed
    blinky.Name = "Blinky"
    Set blinky.CurrentTile = MapManager.GetMazeTile(22, 30)
    Set blinky.ActiveBehavior = ChasePacmanBehavior.Create(player)
    this.Moveables.Add blinky
    UIAdapter.AsCommandSender.CreateGhost blinky
    
    Set inky = New Ghost
    inky.Color = vbCyan
    inky.Name = "Inky"
    Set inky.CurrentTile = MapManager.GetMazeTile(22, 20)
    Set inky.ActiveBehavior = BashfulBehavior.Create(player, blinky)
    this.Moveables.Add inky
    UIAdapter.AsCommandSender.CreateGhost inky
    
    UIAdapter.AsCommandSender.CreateMap this.Maze
    
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
            player.Heading = dUp
        Case KeyCode.DownArrow
            player.Heading = dDown
        Case KeyCode.LeftArrow
            player.Heading = dLeft
        Case KeyCode.RightArrow
            player.Heading = dRight
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


