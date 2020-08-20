VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pellet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine")
Option Explicit
Implements IGamePiece
Implements IDrawable

Public WithEvents HostCell As GridCell
Public Parent As Collection


Private Sub HostCell_DispatchEncounter()
    '// remove the pellet
    HostCell.HostRange.value = vbNullString
     
    '//Decrement pellets from the game and apply score
    If Not Parent Is Nothing Then
        Parent.Remove (Me.ToString)
        
    End If
End Sub

Private Sub IDrawable_Render()
    With HostCell.HostRange
        .value = Drawing.TextDot
        .Font.Size = 5
    End With
End Sub

Private Property Set IGamePiece_HostCell(ByVal RHS As GridCell)
    Set HostCell = RHS
End Property

Private Property Get IGamePiece_HostCell() As GridCell
    Set IGamePiece_HostCell = HostCell
End Property

Public Function ToString() As String
    ToString = HostCell.HostRange.Address
End Function
