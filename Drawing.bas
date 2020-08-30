Attribute VB_Name = "Drawing"
'@Folder("PacManEngine")
Option Explicit

Private Const TEXT_DOT_CHAR_CODE As Integer = 149

Public Const clBLUE As Long = &HFF0000

Public Property Get TextDot() As String
    TextDot = Chr(TEXT_DOT_CHAR_CODE)
End Property

Public Sub CenterShapeOnRange(tarShape As Shape, tarRng As Range)
    tarShape.Left = tarRng.Left - (tarShape.Width / 2) + (tarRng.Width / 2)
    tarShape.Top = tarRng.Top - (tarShape.Height / 2) + (tarRng.Height / 2)
End Sub

Public Function ColorAsRGB(colorCode As Long) As Variant
    ColorAsRGB = Split((colorCode Mod 256) & ", " & ((colorCode \ 256) Mod 256) & ", " & (colorCode \ 65536), ",")
End Function

