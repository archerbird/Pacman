Attribute VB_Name = "Module1"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 384.75, 290.25, 456.75, _
        362.25).Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes( _
        "Rectangle 145"), 2
    Selection.ShapeRange.ScaleWidth 1.3541666667, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.6458333333, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.Flip msoFlipVertical
    Selection.ShapeRange.ScaleWidth 2, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleWidth 0.5, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleWidth 0.0405960946, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 1.8192936701, msoFalse, _
        msoScaleFromBottomRight
    Selection.ShapeRange.Flip msoFlipHorizontal
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Selection.ShapeRange.ScaleWidth 0.9056603774, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.7745098039, msoFalse, _
        msoScaleFromBottomRight
    Selection.ShapeRange.Flip msoFlipVertical
    Selection.ShapeRange.ScaleWidth 0.9791667122, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 1.417721519, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.Flip msoFlipVertical
    Selection.ShapeRange.ScaleWidth 3.170198603, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.3125002929, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.Flip msoFlipVertical
    Selection.ShapeRange.ScaleWidth 0.7382548491, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 3.8857142857, msoFalse, _
        msoScaleFromBottomRight
    Selection.ShapeRange.Flip msoFlipVertical
End Sub
