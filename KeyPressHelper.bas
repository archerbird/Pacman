Attribute VB_Name = "KeyPressHelper"
Option Explicit
Public userInput As KeyPressDispatcher

Public Sub KeyPressed(keyCode As KeyCodes)
    If userInput Is Nothing Then
        '//get rid of the OnKey hooks
        Set userInput = New KeyPressDispatcher
        Set userInput = Nothing
        Exit Sub
    End If
    
    userInput.KeyPressed keyCode
    
End Sub

Public Sub LeftArrowPressed()
    KeyPressed LeftArrow
 End Sub
Public Sub RightArrowPressed()
    KeyPressed RightArrow
End Sub
Public Sub UpArrowPressed()
    KeyPressed UpArrow
End Sub
Public Sub DownArrowPressed()
    KeyPressed DownArrow
End Sub
