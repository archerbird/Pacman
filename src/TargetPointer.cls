VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TargetPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ViewImplementations.ExcelWorksheet.ShapeWrapper")
Option Explicit
Implements IDrawable

Private Type TTargetPointer
    InternalShape As Shape
    HostGhost As GhostStyler
    ArrowHead As Shape
End Type

Private Const clBLUE As Long = &HC47244
Private Const clWHITE As Long = &HFFFFFF

Private this As TTargetPointer

Public Property Get AsShape() As Shape
    Set AsShape = this.InternalShape
End Property

Public Sub Init(hostingGhost As GhostStyler)
    Dim wks As Worksheet
    Dim pntr As Shape
    Dim head As Shape
    Set wks = hostingGhost.AsShape.Parent
    
    Set head = wks.Shapes.AddShape(msoShapeOval, 0, 0, 1, 1)
    Set this.ArrowHead = head
    
    Set pntr = wks.Shapes.AddConnector(msoConnectorStraight, 0, 0, 0, 0)
    pntr.Line.EndArrowheadStyle = msoArrowheadTriangle
    pntr.Line.ForeColor.RGB = hostingGhost.AsShape.GroupItems(1).Fill.ForeColor
    pntr.ConnectorFormat.BeginConnect hostingGhost.AsShape.GroupItems(1), 1
    pntr.ConnectorFormat.EndConnect head, 1
    Set this.InternalShape = pntr
    
End Sub

Private Sub IDrawable_Redraw(facing As Direction, location As Range)

    Drawing.CenterShapeOnRange this.ArrowHead, location
    
End Sub

