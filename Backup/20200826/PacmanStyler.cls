VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PacmanStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ViewImplementations.ExcelWorksheet"
Option Explicit

Private mSelf As Shape
Private mMoveInterval As Long

'@DefaultMember
Public Property Get Self() As Shape
Attribute Self.VB_UserMemId = 0
    Set Self = mSelf
End Property

Public Sub Redraw(facing As Direction, location As Range)
    Select Case facing
        Case Direction.dLeft
            Self.Rotation = 0
            
        Case Direction.dUp
            Self.Rotation = 90
        Case Direction.dRight
            Self.Rotation = 180
        Case Direction.dDown
            Self.Rotation = 270
    End Select
    
    Drawing.CenterShapeOnRange Me.Self, location
End Sub


Public Function CreateShape(hostSheet As Worksheet) As Shape
    Dim newShape As Shape
    
    Set newShape = hostSheet.Shapes.AddShape(msoShapePie, 0, 0, 20.26, 20.26)
    newShape.Flip msoFlipHorizontal
    newShape.Adjustments(1) = 35
    newShape.Adjustments(2) = -35
    newShape.Line.Visible = msoFalse
    newShape.Fill.ForeColor.RGB = vbYellow
    
    Set mSelf = newShape
    Set CreateShape = mSelf
End Function
