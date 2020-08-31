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

Private mIsDecisionNodeSupressed As Boolean
Private mAddress As Coord

Public Event DispatchEncounter()

Public Property Get IsDecisionNode() As Boolean
    '// true if there are 3 or more path options
    If mIsDecisionNodeSupressed = True Then Exit Function
    IsDecisionNode = -1 * (MapManager.GetNextTile(Me, dDown).IsTraversable + _
                            MapManager.GetNextTile(Me, dUp).IsTraversable + _
                            MapManager.GetNextTile(Me, dRight).IsTraversable + _
                            MapManager.GetNextTile(Me, dLeft).IsTraversable) _
                        >= 3
End Property
Public Property Let IsDecisionNode(value As Boolean)
    mIsDecisionNodeSupressed = value
End Property

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
