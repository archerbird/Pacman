VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameUICommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI")
'@Interface
'// Commands that the Controller will need to be able to send to the UI
Option Explicit
Public Property Get Events() As GameUIAdapter
End Property

Public Property Set Events(ByVal value As GameUIAdapter)
End Property
Public Sub DrawMap(map() As Tile)
End Sub

Public Sub DrawPacman()
End Sub

Public Sub DrawGhost()
End Sub

Private Sub Class_Initialize()
    '//Err.Raise 5, mModuleName, AccessError(5) & "-Interface class must not be instantiated."
End Sub
