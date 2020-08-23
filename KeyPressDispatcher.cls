VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyPressDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI")
'//Concrete implementation for events
Option Explicit
Implements IGameUIEvents

Public Enum KeyCodes
    LeftArrow = 37
    RightArrow = 39
    UpArrow = 38
    DownArrow = 40
End Enum

Private Const LEFT_ARROW_KEY   As String = "{LEFT}"
Private Const RIGHT_ARROW_KEY  As String = "{RIGHT}"
Private Const DOWN_ARROW_KEY   As String = "{DOWN}"
Private Const UP_ARROW_KEY  As String = "{UP}"


Public Event KeyPressUp()
Public Event KeyPressDown()
Public Event KeyPressLeft()
Public Event KeyPressRight()
Public Event GameStarted()


Public Sub Attach()
    '// point the Application.OnKeys to the exposed versions of the key press.
    Set KeyPressHelper.userInput = Me
    Application.OnKey LEFT_ARROW_KEY, "LeftArrowPressed"
    Application.OnKey RIGHT_ARROW_KEY, "RightArrowPressed"
    Application.OnKey DOWN_ARROW_KEY, "DownArrowPressed"
    Application.OnKey UP_ARROW_KEY, "UpArrowPressed"
End Sub

Public Sub Detach()
    '// point the Application.OnKeys back to normal
    Set KeyPressHelper.userInput = Nothing
    Application.OnKey LEFT_ARROW_KEY
    Application.OnKey RIGHT_ARROW_KEY
    Application.OnKey DOWN_ARROW_KEY
    Application.OnKey UP_ARROW_KEY
End Sub
Public Sub KeyPressed(keyCode As KeyCodes)
    Select Case keyCode
        Case KeyCodes.DownArrow
            IGameUIEvents_OnDownKey
        Case KeyCodes.UpArrow
            RaiseEvent KeyPressUp
        Case KeyCodes.LeftArrow
           IGameUIEvents_OnLeftkey
        Case KeyCodes.RightArrow
            IGameUIEvents_OnRightKey
    End Select
End Sub

Private Sub Class_Terminate()
    Debug.Print TypeName(Me) & " terminating..."
End Sub

Private Sub IGameUIEvents_OnDownKey()
    RaiseEvent KeyPressDown
End Sub

Private Sub IGameUIEvents_OnGameStarted()
    RaiseEvent GameStarted
End Sub

Private Sub IGameUIEvents_OnLeftkey()
    RaiseEvent KeyPressLeft
End Sub

Private Sub IGameUIEvents_OnRightKey()
    RaiseEvent KeyPressRight
End Sub

Private Sub IGameUIEvents_OnUpKey()
    RaiseEvent KeyPressUp
End Sub
