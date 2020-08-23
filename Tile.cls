VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.GridCells")
Option Explicit

Public Event DispatchEncounter()

Public IsTraversable As Boolean
Public Id As String


Public Sub TriggerEncounterInteractions()
'//This is called by PacMan whenever he enters a grid
'//Any other GamePieces will here this event and respond accordingly
    RaiseEvent DispatchEncounter
End Sub

Private Sub Class_Terminate()
    'Debug.Print TypeName(Me) & " Id:" & Me.Id & " terminating..."
End Sub
