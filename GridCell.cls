VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.GridCells")
Option Explicit
Implements IDrawable

Public Enum WallAlignment
    Horizontal = 1
    Vertical = 2
End Enum

Private mIsDecomposing As Boolean

Public LeftPathNode As GridCell
Public RightPathNode As GridCell
Public UpperPathNode As GridCell
Public LowerPathNode As GridCell

Public HostRange As Range
Public Pen As Line
Public IsTraversable As Boolean
Public Alignment As WallAlignment

Public Event DispatchEncounter()

Public Property Get Id() As String
    If Not HostRange Is Nothing Then
        Id = HostRange.Address
    End If
End Property

Private Sub IDrawable_Render()
    If Pen Is Nothing Then
        Exit Sub
    End If
    
    If Me.IsTraversable Then
        
    Else
        If LeftPathNode.IsTraversable Then
            Pen.Draw HostRange.Borders(xlEdgeRight)
        End If
        
        If LowerPathNode.IsTraversable Then
            Pen.Draw HostRange.Borders(xlEdgeTop)
        End If
        
        If RightPathNode.IsTraversable Then
            Pen.Draw HostRange.Borders(xlEdgeLeft)
        End If
        
        If UpperPathNode.IsTraversable Then
            Pen.Draw HostRange.Borders(xlEdgeBottom)
        End If
    End If
End Sub

Public Sub Decompose()
    If mIsDecomposing Then
        Exit Sub
    End If
    
    mIsDecomposing = True
    
    If Not LeftPathNode Is Nothing Then
        LeftPathNode.Decompose
        Set LeftPathNode = Nothing
    End If
    
    If Not RightPathNode Is Nothing Then
        RightPathNode.Decompose
        Set RightPathNode = Nothing
    End If
    
    If Not UpperPathNode Is Nothing Then
        UpperPathNode.Decompose
        Set UpperPathNode = Nothing
    End If
    
    If Not LowerPathNode Is Nothing Then
        LowerPathNode.Decompose
        Set LowerPathNode = Nothing
    End If
End Sub

Public Sub TriggerEncounterInteractions()
'//This is called by PacMan whenever he enters a grid
'//Any other GamePieces will here this event and respond accordingly
    RaiseEvent DispatchEncounter
End Sub

