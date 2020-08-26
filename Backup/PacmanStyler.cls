VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PacmanStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
Option Explicit

Public HostSheet As Worksheet
Private WithEvents mAbstractPacman As PacmanController
Private mSelf As Shape
Private mMoveInterval As Long

'@DefaultMember
Public Property Get Self() As Shape
Attribute Self.VB_UserMemId = 0
    Set Self = mSelf
End Property

Private Sub Redraw()
    Select Case mAbstractPacman.Heading
        Case Direction.dLeft
            Self.Rotation = 0
            
        Case Direction.dUp
            Self.Rotation = 90
        Case Direction.dRight
            Self.Rotation = 180
        Case Direction.dDown
            Self.Rotation = 270
    End Select
End Sub


Public Function CreateShape() As Shape
    Dim newShape As Shape
    
    Set newShape = HostSheet.Shapes.AddShape(msoShapePie, 0, 0, 20.26, 20.26)
    newShape.Flip msoFlipHorizontal
    newShape.Adjustments(1) = 35
    newShape.Adjustments(2) = -35
    newShape.Line.Visible = msoFalse
    
    Set mSelf = newShape
    Set CreateShape = Me
End Function
