VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IViewEventsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Methods that the UI can call to notify the controller of user interaction. These are events from the UI that the Controller wants to hear about"
'@Folder "PacmanGame.View"
'@Interface
'@ModuleDescription("Methods that the UI can call to notify the controller of user interaction. These are events from the UI that the Controller wants to hear about")

Option Explicit
Private Const mModuleName As String = "IViewEventsHandler"

Public Enum KeyCode
    LeftArrow = 37
    RightArrow = 39
    UpArrow = 38
    DownArrow = 40
End Enum

Public Sub OnDirectionalKeyPress(vbKey As KeyCode)
End Sub

Public Sub OnGameStarted()
End Sub

Public Sub OnGamePaused()
End Sub

Public Sub OnQuit()
End Sub

Private Sub Class_Initialize()
    Err.Raise 5, mModuleName, "Interface class must not be instantiated."
End Sub
