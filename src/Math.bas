Attribute VB_Name = "Math"
'@Folder "PacmanGame.Common"
Option Explicit

Public Function min(x As Long, y As Long) As Long
   min = IIf(x < y, x, y)
End Function
