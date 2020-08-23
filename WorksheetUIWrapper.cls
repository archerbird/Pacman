VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetUIWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
'//UI implemented as an Excel Worksheet
Option Explicit
Implements IGameUICommands

Private Const MAP_ADDRESS As String = "$C$2:$BE$62"
Private WithEvents xlWs As Worksheet
Attribute xlWs.VB_VarHelpID = -1
Private mAdapter As GameUIAdapter

Public Event Activated()
Public Event Deactivated()


Public Sub Init(xlWs As Worksheet)
    Set xlWs = xlWs
End Sub

Private Sub Class_Terminate()
    Set mAdapter = Nothing
    Set xlWs = Nothing
     Debug.Print TypeName(Me) & " terminating..."
End Sub

'// Support for IGameUICommands
Private Sub IGameUICommands_DrawGhost()

End Sub

Private Sub IGameUICommands_DrawMap(map() As Tile)

End Sub

Private Sub IGameUICommands_DrawPacman()

End Sub

Private Property Set IGameUICommands_Events(ByVal RHS As GameUIAdapter)
    Set mAdapter = RHS
End Property

Private Property Get IGameUICommands_Events() As GameUIAdapter
    Set IGameUICommands_Events = mAdapter
End Property

'//Adds/Removes Keyboard Listening when the sheet does active/inactive
Private Sub xlWs_Activate()
    RaiseEvent Activated
End Sub

Private Sub xlWs_Deactivate()
    RaiseEvent Deactivated
End Sub
