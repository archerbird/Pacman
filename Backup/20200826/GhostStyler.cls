VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GhostStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ViewImplementations.ExcelWorksheet"
Option Explicit

Private mSelf As Shape
Private Const clBLUE As Long = &HC47244
Private Const clWHITE As Long = &HFFFFFF

Public Property Get Self() As Shape
    Set Self = mSelf
End Property

Public Sub Redraw(facing As Direction, location As Range)

     Drawing.CenterShapeOnRange Me.Self, location
End Sub
Public Function CreateShape(hostSheet As Worksheet, ghostColor As Long) As Shape
    Dim newShapesNames(0 To 7) As String
    Dim newShapeNamesEyes(0 To 4) As String
    Dim result As Shape
    Dim shapePointer As Shape
    Dim ghostBody As Shape
    With hostSheet.Shapes
        Set shapePointer = .AddShape(msoShapeRoundedRectangle, 0, 0, 20.65, 15.15) '//Head
        shapePointer.Adjustments(1) = 0.5
        newShapesNames(0) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRectangle, 0, 7.37, 20.65, 8.4) '//Body
        newShapesNames(1) = shapePointer.Name
        
        '//Bottom wisps
        Set shapePointer = .AddShape(msoShapeRightTriangle, 17.063, 15.75, 3.6, 3.6) '//LeftMost
        shapePointer.Flip msoFlipVertical
        newShapesNames(2) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRightTriangle, 13.5, 15.75, 3.6, 3.6) '//CenterLeft
        shapePointer.Flip msoFlipVertical
        shapePointer.Flip msoFlipHorizontal
        newShapesNames(3) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRightTriangle, 3.56, 15.75, 3.6, 3.6) '//CenterRight
        shapePointer.Flip msoFlipVertical
        newShapesNames(4) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRightTriangle, 0, 15.75, 3.6, 3.6) 'RightMost
        shapePointer.Flip msoFlipVertical
        shapePointer.Flip msoFlipHorizontal
        newShapesNames(5) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRightTriangle, 10.3, 15.75, 3.6, 3.6) 'RightMost
        shapePointer.Flip msoFlipVertical
        newShapesNames(6) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeRightTriangle, 6.75, 15.75, 3.6, 3.6) 'RightMost
        shapePointer.Flip msoFlipVertical
        shapePointer.Flip msoFlipHorizontal
        newShapesNames(7) = shapePointer.Name
        
        Set ghostBody = hostSheet.Shapes.Range(newShapesNames).Group
        ghostBody.Fill.ForeColor.RGB = ghostColor
        'ghostBody.Line.Visible = msoFalse
        newShapeNamesEyes(0) = ghostBody.Name
        
        '//Eyes
        Set shapePointer = .AddShape(msoShapeOval, 0.47, 3.8, 6#, 5.5)
        shapePointer.Fill.ForeColor.RGB = clWHITE
        newShapeNamesEyes(1) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeOval, 7.26, 3.8, 6#, 5.5)
        shapePointer.Fill.ForeColor.RGB = clWHITE
        newShapeNamesEyes(2) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeOval, 0.2, 4.6, 3.6, 3.6)
        shapePointer.Fill.ForeColor.RGB = clBLUE
        newShapeNamesEyes(3) = shapePointer.Name
        
        Set shapePointer = .AddShape(msoShapeOval, 7#, 4.6, 3.6, 3.6)
        shapePointer.Fill.ForeColor.RGB = clBLUE
        newShapeNamesEyes(4) = shapePointer.Name
        End With
        
    Set result = hostSheet.Shapes.Range(newShapeNamesEyes).Group
    result.Height = 20.6
    result.Width = 20.6
    
    result.Line.Visible = msoFalse
    Set mSelf = result
    Set CreateShape = result
    
End Function

