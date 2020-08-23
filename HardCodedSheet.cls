VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HardCodedSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
Option Explicit
Implements IGameUICommands
Implements WorksheetUIWrapper

Private Const MAP_ADDRESS As String = "$D$3:$BD$61"
Private WithEvents sheetUI As PacmanUI
Attribute sheetUI.VB_VarHelpID = -1
Private mAdapter As GameUIAdapter

Public Event Activated()
Public Event Deactivated()


Private Sub Class_Initialize()
    Set sheetUI = PacmanUI
End Sub

Private Sub Class_Terminate()
    Set sheetUI = Nothing
    Set mAdapter = Nothing
    Debug.Print TypeName(Me) & " terminating..."
End Sub

Private Sub IGameUICommands_DrawGhost()

End Sub

Private Sub IGameUICommands_DrawMap(map() As Tile)
    '//this hardcoded version comes with the map already drawn
    
End Sub

Private Sub IGameUICommands_DrawPacman()

End Sub

Private Property Set IGameUICommands_Events(ByVal RHS As GameUIAdapter)
    Set mAdapter = RHS
End Property

Private Property Get IGameUICommands_Events() As GameUIAdapter
    Set IGameUICommands_Events = mAdapter
End Property

Private Sub sheetUI_Activated()
    RaiseEvent Activated
End Sub

Private Sub sheetUI_Deactivated()
    RaiseEvent Deactivated
End Sub

Private Sub WorksheetUIWrapper_Init(xlWs As Worksheet)

End Sub
