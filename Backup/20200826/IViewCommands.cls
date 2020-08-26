VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IViewCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.View"
'@Interface
'// Commands that the Controller will need to be able to send to the UI
Option Explicit
Public Property Get Events() As ViewAdapter
End Property

Public Property Set Events(ByVal value As ViewAdapter)
End Property
Public Sub DrawMap(map() As Tile)
End Sub

Public Sub DrawPacman(character As PacmanController)
End Sub

Public Sub DrawGhost(character As Ghost)
End Sub

Private Sub Class_Initialize()
    '//Err.Raise 5, mModuleName, AccessError(5) & "-Interface class must not be instantiated."
End Sub
