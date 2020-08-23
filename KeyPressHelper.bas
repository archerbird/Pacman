Attribute VB_Name = "KeyPressHelper"
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
Option Explicit
Public userInput As KeyPressDispatcher
'// This module is only here to expose the class methods to the Application.OnKey Action pointer

Public Sub KeyPressed(keyCode As KeyCodes)
    If userInput Is Nothing Then
        '//get rid of the OnKey hooks
        Set userInput = New KeyPressDispatcher
        userInput.Detach
        Set userInput = Nothing
        Exit Sub
    End If
    
    userInput.KeyPressed keyCode
    
End Sub

'// these methods delegate back into the class
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
