VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyPressDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine")
Option Explicit

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


Private Sub Class_Initialize()
    Application.OnKey LEFT_ARROW_KEY, "LeftArrowPressed"
    Application.OnKey RIGHT_ARROW_KEY, "RightArrowPressed"
    Application.OnKey DOWN_ARROW_KEY, "DownArrowPressed"
    Application.OnKey UP_ARROW_KEY, "UpArrowPressed"
End Sub
    
Private Sub Class_Terminate()
    Application.OnKey LEFT_ARROW_KEY
    Application.OnKey RIGHT_ARROW_KEY
    Application.OnKey DOWN_ARROW_KEY
    Application.OnKey UP_ARROW_KEY
End Sub

Public Sub KeyPressed(keyCode As KeyCodes)
    Select Case keyCode
        Case KeyCodes.DownArrow
            RaiseEvent KeyPressDown
        Case KeyCodes.UpArrow
            RaiseEvent KeyPressUp
        Case KeyCodes.LeftArrow
            RaiseEvent KeyPressLeft
        Case KeyCodes.RightArrow
            RaiseEvent KeyPressRight
    End Select
End Sub
