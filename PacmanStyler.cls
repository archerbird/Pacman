VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PacmanStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "ViewImplementations.ExcelWorksheet.ShapeWrapper"
Option Explicit
Implements IDrawable

Private Type TPacmanViewModel
    InternalShape As Shape
End Type

Private this As TPacmanViewModel

'@DefaultMember
Public Property Get AsShape() As Shape
Attribute AsShape.VB_UserMemId = 0
    Set AsShape = this.InternalShape
End Property

Public Function Create(hostSheet As Worksheet) As IDrawable
    Dim result As New PacmanStyler
    
    result.Init hostSheet
    
    Set Create = result
End Function

Public Sub Init(hostSheet As Worksheet)
    Dim newShape As Shape
    
    
    Set newShape = hostSheet.Shapes.AddShape(msoShapePie, 0, 0, 20.26, 20.26)
    With newShape
        .Flip msoFlipHorizontal
        .Adjustments(1) = 35
        .Adjustments(2) = -35
        .Line.Visible = msoFalse
        .Fill.ForeColor.RGB = vbYellow
    End With
    
    Set this.InternalShape = newShape
End Sub

Private Sub IDrawable_Redraw(facing As Direction, location As Range)
    Select Case facing
        Case Direction.dLeft
            AsShape.Rotation = 0
        Case Direction.dUp
            AsShape.Rotation = 90
        Case Direction.dRight
            AsShape.Rotation = 180
        Case Direction.dDown
            AsShape.Rotation = 270
    End Select
    
    Drawing.CenterShapeOnRange Me.AsShape, location
End Sub
