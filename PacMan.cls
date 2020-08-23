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

Public HostCell As Tile

Private Property Set IGamePiece_HostCell(ByVal RHS As Tile)
        Set HostCell = RHS
        HostCell.TriggerEncounterInteractions
End Property

Private Property Get IGamePiece_HostCell() As Tile
    Set IGamePiece_HostCell = HostCell
End Property


