VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameUIEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI")
'@Interface
'// Events from the UI that the Controller wants to hear
Option Explicit

Private Sub Class_Initialize()
    Err.Raise 5, , "-Interface class must not be instantiated."
End Sub

Public Sub OnUpKey()
End Sub

Public Sub OnDownKey()
End Sub

Public Sub OnLeftkey()
End Sub

Public Sub OnRightKey()
End Sub

Public Sub OnGameStarted()
End Sub
