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
Private Const mModuleName As String = "IViewCommands"

'@Description("Provides a way for the ViewAdapter to hook itself into an IViewCommands implementer")
Public Property Get Events() As IViewEvents
Attribute Events.VB_Description = "Provides a way for the ViewAdapter to hook itself into an IViewCommands implementer"
End Property
Public Property Set Events(ByVal value As IViewEvents)
End Property

Public Sub CreateMap(map() As Tile)
End Sub

Public Sub CreatePacman(character As PacmanController)
End Sub

Public Sub CreateGhost(character As Ghost)
End Sub

Public Sub UpdateComponent(characters As Collection)
End Sub

Private Sub Class_Initialize()
    Err.Raise 5, mModuleName, "Interface class must not be instantiated."
End Sub
