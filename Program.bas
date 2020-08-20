Attribute VB_Name = "Program"
'@Folder("PacManEngine")
Option Explicit
Private mController As GameController


Public Sub Main()
    Set KeyPressHelper.userInput = New KeyPressDispatcher
    Set mController = New GameController
    
    mController.StartGame KeyPressHelper.userInput
End Sub

