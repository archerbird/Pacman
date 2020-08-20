Attribute VB_Name = "Drawing"
'@Folder("PacManEngine.GridCells")
Option Explicit

Private Const TEXT_DOT_CHAR_CODE As Integer = 149

Public Const clBLUE As Long = &HFF0000

Public Property Get TextDot() As String
    TextDot = Chr(TEXT_DOT_CHAR_CODE)
End Property

Public Function OutsideBorderLine() As Line
    Set OutsideBorderLine = New Line
    With OutsideBorderLine
        .Color = clBLUE
        .Style = xlDouble
        .weight = xlThick
    End With
  
End Function

Public Function InsideBorderLine() As Line
    Set OutsideBorderLine = New Line
    With OutsideBorderLine
        .Color = clBLUE
        .Style = xlContinuous
        .weight = xlMedium
    End With
  
End Function

Public Sub Render(obj As IDrawable)
    obj.Render
End Sub

