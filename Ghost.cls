VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ghost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit
Implements IGamePiece
Implements IDrawable

Public WithEvents HostCell As GridCell
Attribute HostCell.VB_VarHelpID = -1

Private Sub HostCell_DispatchEncounter()
    '//You have encounted Pacman!
    
End Sub

Private Sub IDrawable_Render()

End Sub

Private Property Set IGamePiece_HostCell(ByVal RHS As GridCell)
    Set HostCell = RHS
End Property

Private Property Get IGamePiece_HostCell() As GridCell
    Set IGamePiece_HostCell = HostCell
End Property

