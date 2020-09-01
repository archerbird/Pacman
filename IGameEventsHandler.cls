VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGameEventsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Methods that the Controller will need to be able to call in the UI. These are things the Controller will need to tell the UI to do."
'@Folder "PacmanGame.View"
'@Interface
'@ModuleDescription("Methods that the Controller will need to be able to call in the UI. These are things the Controller will need to tell the UI to do.")
Option Explicit
Private Const mModuleName As String = "IGameEventsHandler"

'@Description("Provides a way for the ViewAdapter to hook itself into an  IGameEventsHandler implementer")
Public Property Get Events() As IViewEventsHandler
Attribute Events.VB_Description = "Provides a way for the ViewAdapter to hook itself into an  IGameEventsHandler implementer"
End Property
Public Property Set Events(ByVal value As IViewEventsHandler)
End Property

Public Sub CreateMap(map() As Tile)
End Sub

Public Sub CreatePacman(character As PacmanModel)
End Sub

Public Sub CreateGhost(character As GhostModel)
End Sub

Public Sub UpdateComponents(GamePieces As Collection)
End Sub

Private Sub Class_Initialize()
    Err.Raise 5, mModuleName, "Interface class must not be instantiated."
End Sub
