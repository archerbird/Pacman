VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.Model.Maps"
Option Explicit
Private Type Coord
    y As Integer
    x As Integer
End Type

Public IsTraversable As Boolean
Public Id As String
Private mAddress As Coord

Public Event DispatchEncounter()

Public Property Get x() As Integer
    x = mAddress.x
End Property
Public Property Get y() As Integer
    y = mAddress.y
End Property

Public Sub SetAddress(xPos As Integer, yPos As Integer)
    mAddress.x = xPos
    mAddress.y = yPos
End Sub

Public Sub TriggerEncounterInteractions()
'//This is called by PacMan whenever he enters a grid
'//Any other GamePieces will here this event and respond accordingly
    RaiseEvent DispatchEncounter
End Sub

Private Sub Class_Terminate()
    'Debug.Print TypeName(Me) & " Id:" & Me.Id & " terminating..."
End Sub
