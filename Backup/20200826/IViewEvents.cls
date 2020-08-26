VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IViewEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.View"
'@Interface
'// Events from the UI that the Controller wants to hear
Option Explicit

Private Sub Class_Initialize()
    Err.Raise 5, , "-Interface class must not be instantiated."
End Sub

Public Sub OnDirectionalKeyPress(vbKey As KeyCode)
End Sub

Public Sub OnGameStarted()
End Sub

