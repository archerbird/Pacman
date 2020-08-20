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

Private mBoard As GameGrid
Private WithEvents mUiHost As Worksheet
Attribute mUiHost.VB_VarHelpID = -1
Private WithEvents mControls As KeyPressDispatcher
Attribute mControls.VB_VarHelpID = -1


Private Sub BuildGameBoard()

End Sub

Private Sub PointPacMan(pathing As KeyCodes)
    Debug.Print "PacMan moved " & pathing
End Sub

Private Sub Class_Initialize()
    Set mUiHost = ThisWorkbook.Worksheets("PacmanUI")
    BuildGameBoard
End Sub

Private Sub Class_Terminate()
    
    Set mUiHost = Nothing
End Sub

Public Sub StartGame(userInput As KeyPressDispatcher)
    Set mControls = userInput
End Sub

Private Sub mControls_KeyPressDown()
    PointPacMan KeyCodes.DownArrow
End Sub

Private Sub mControls_KeyPressLeft()
    PointPacMan KeyCodes.LeftArrow
End Sub

Private Sub mControls_KeyPressRight()
    PointPacMan KeyCodes.RightArrow
End Sub

Private Sub mControls_KeyPressUp()
    PointPacMan KeyCodes.UpArrow
End Sub

Private Sub mUiHost_Activate()
    Set mControls = New KeyPressDispatcher
End Sub

Private Sub mUiHost_Deactivate()
    Set mControls = Nothing
End Sub

