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
Private WithEvents innerWs As Worksheet
Attribute innerWs.VB_VarHelpID = -1
Private mAdapter As GameUIAdapter
Private dPad As Range

Public Enum KeyCode
    LeftArrow = 37
    RightArrow = 39
    UpArrow = 38
    DownArrow = 40
End Enum

Public Event Activated()
Public Event Deactivated()
Public Event DirectionalKeyPressed(vbKey As KeyCode)


Public Sub Init(xlWs As Worksheet)
    
    xlWs.Activate
    xlWs.Range("A65").Select
    Set innerWs = xlWs
    Set dPad = xlWs.Range("A65")
End Sub

Private Sub Class_Terminate()
    Set mAdapter = Nothing
    Set innerWs = Nothing
    Set dPad = Nothing
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
Private Sub innerWs_Activate()
    RaiseEvent Activated
End Sub

Private Sub innerWs_Deactivate()
    RaiseEvent Deactivated
End Sub

Private Sub innerWs_SelectionChange(ByVal Target As Range)
    If dPad.Offset(-1, 0).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(DownArrow)
    ElseIf dPad.Offset(1, 0).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(UpArrow)
    ElseIf dPad.Offset(0, -1).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(LeftArrow)
    ElseIf dPad.Offset(0, 1).Address = Target.Address Then
        RaiseEvent DirectionalKeyPressed(RightArrow)
    End If
    
    Application.EnableEvents = False
    innerWs.Range("A65").Select
    Application.EnableEvents = True
End Sub
