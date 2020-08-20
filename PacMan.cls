VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PacMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit
Implements IGamePiece

Public HostCell As GridCell

Private Property Set IGamePiece_HostCell(ByVal RHS As GridCell)
        Set HostCell = RHS
        HostCell.TriggerEncounterInteractions
End Property

Private Property Get IGamePiece_HostCell() As GridCell
    Set IGamePiece_HostCell = HostCell
End Property


